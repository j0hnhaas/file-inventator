#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
  Stage 3: copy an approved Recovery Audit Toolkit sample plan to a safe
  destination with progress reporting and an audit manifest.

.DESCRIPTION
  recovery-audit-copy.ps1 consumes sample-plan.csv produced by
  recovery-audit-sampler.ps1.

  It does not choose files. It copies exactly the rows in the approved plan,
  preserving each row's RelativePath below DestinationRoot.

  Safety controls:
    - source disk serial must match ExpectedSerial;
    - Windows must report the source disk as read-only;
    - read-only state is checked before the run, every 30 seconds, and at end;
    - the plan and destination must be on a different volume from the source;
    - destination paths are canonicalised and must stay below DestinationRoot;
    - duplicate/colliding destination paths abort preflight;
    - existing non-empty destination folders are rejected;
    - files are never overwritten;
    - copied file length is verified against SizeBytes from the plan.

  Progress is shown by processed bytes and includes file count, byte count,
  throughput, ETA, and current source path.

  The script calculates no source-file hashes.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-copy.ps1" -SamplePlan ".\sample-plan.csv" -SourceRoot "X:\" -ExpectedSerial "SERIAL_NUMBER" -DestinationRoot "C:\RecoveryAuditSample" -ExpectedPlanFiles 1000 -ExpectedPlanBytes 3900000000
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$SamplePlan,

  [Parameter(Mandatory=$true)]
  [ValidatePattern('^[A-Za-z]:\\?$')]
  [string]$SourceRoot,

  [Parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$ExpectedSerial,

  [Parameter(Mandatory=$true)]
  [string]$DestinationRoot,

  [Parameter()]
  [UInt64]$ExpectedPlanFiles=0,

  [Parameter()]
  [UInt64]$ExpectedPlanBytes=0,

  [Parameter()]
  [UInt64]$MinimumFreeReserveBytes=100000000
)

$ErrorActionPreference='Stop'
$ScriptVersion='1.0.1'
$RunStarted=Get-Date

function Normalize-Root([string]$Path) {
  $p=[System.IO.Path]::GetFullPath($Path)
  if(-not $p.EndsWith('\')){$p+='\'}
  return $p
}

function Csv([AllowNull()][string]$Value) {
  if($null -eq $Value){return '""'}
  return '"' + $Value.Replace('"','""') + '"'
}

function Format-Duration([TimeSpan]$Span) {
  $h=[int][math]::Floor($Span.TotalHours)
  return ('{0:00}:{1:00}:{2:00}' -f $h,$Span.Minutes,$Span.Seconds)
}

function Format-DecimalGB([UInt64]$Bytes) {
  return ([math]::Round(([double]$Bytes/1000000000),3))
}

function Get-CheckedDisk([string]$Root,[string]$Serial) {
  $letter=$Root.Substring(0,1)
  $partition=Get-Partition -DriveLetter $letter -ErrorAction Stop
  $disk=$partition | Get-Disk -ErrorAction Stop
  $actual=([string]$disk.SerialNumber).Trim()

  if($actual -ne $Serial.Trim()){
    throw "SAFETY ABORT: source serial '$actual' does not match expected serial '$Serial'."
  }

  if(-not $disk.IsReadOnly){
    throw "SAFETY ABORT: source disk $($disk.Number) is NOT read-only."
  }

  return $disk
}

function Assert-ReadOnly([int]$DiskNumber,[string]$Serial) {
  $disk=Get-Disk -Number $DiskNumber -ErrorAction Stop

  if((([string]$disk.SerialNumber).Trim()) -ne $Serial.Trim()){
    throw 'SAFETY ABORT: source disk identity changed.'
  }

  if(-not $disk.IsReadOnly){
    throw 'SAFETY ABORT: source disk is no longer read-only.'
  }

  return $disk
}

function Convert-ToExtendedPath([string]$Path) {
  $full=[System.IO.Path]::GetFullPath($Path)
  if($full.StartsWith('\\?\')){return $full}
  if($full.StartsWith('\\')){
    return '\\?\UNC\'+$full.Substring(2)
  }
  return '\\?\'+$full
}

function Test-FileExtended([string]$Path) {
  return [System.IO.File]::Exists((Convert-ToExtendedPath $Path))
}

function Get-FileLengthExtended([string]$Path) {
  $extended=Convert-ToExtendedPath $Path
  $stream=New-Object System.IO.FileStream($extended,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
  try{return [UInt64]$stream.Length}
  finally{$stream.Dispose()}
}

function Remove-FileExtended([string]$Path) {
  $extended=Convert-ToExtendedPath $Path
  if([System.IO.File]::Exists($extended)){
    [System.IO.File]::Delete($extended)
  }
}

function Get-CanonicalDestination([string]$Root,[string]$RelativePath) {
  if([string]::IsNullOrWhiteSpace($RelativePath)){
    throw 'RelativePath is empty.'
  }

  $rel=$RelativePath.TrimStart([char[]]@('\','/'))
  $candidate=[System.IO.Path]::GetFullPath((Join-Path $Root $rel))

  if(-not $candidate.StartsWith($Root,[System.StringComparison]::OrdinalIgnoreCase)){
    throw "Destination path escapes DestinationRoot: $RelativePath"
  }

  return $candidate
}

function Copy-FileNoOverwrite([string]$Source,[string]$Destination) {
  $parent=[System.IO.Path]::GetDirectoryName($Destination)
  if([string]::IsNullOrWhiteSpace($parent)){throw "Could not resolve destination directory: $Destination"}

  $sourceExtended=Convert-ToExtendedPath $Source
  $destinationExtended=Convert-ToExtendedPath $Destination
  $parentExtended=Convert-ToExtendedPath $parent

  if(-not [System.IO.Directory]::Exists($parentExtended)){
    [void][System.IO.Directory]::CreateDirectory($parentExtended)
  }

  if([System.IO.File]::Exists($destinationExtended)){
    throw "Destination already exists: $Destination"
  }

  try{
    [System.IO.File]::Copy($sourceExtended,$destinationExtended,$false)
    return
  }
  catch{
    $firstError=$_.Exception.Message

    if([System.IO.File]::Exists($destinationExtended)){
      try{[System.IO.File]::Delete($destinationExtended)}catch{}
    }

    $sourceDir=[System.IO.Path]::GetDirectoryName($Source)
    $fileName=[System.IO.Path]::GetFileName($Source)

    if([string]::IsNullOrWhiteSpace($sourceDir) -or [string]::IsNullOrWhiteSpace($fileName)){
      throw "Primary copy failed and Robocopy fallback could not resolve source path. Primary error: $firstError"
    }

    $robo=Join-Path $env:SystemRoot 'System32\robocopy.exe'
    if(-not(Test-Path -LiteralPath $robo)){
      throw "Primary copy failed and Robocopy is unavailable. Primary error: $firstError"
    }

    & $robo $sourceDir $parent $fileName '/B' '/COPY:DAT' '/DCOPY:T' '/R:0' '/W:0' '/NP' '/NFL' '/NDL' '/NJH' '/NJS' | Out-Null
    $exit=[int]$LASTEXITCODE

    if($exit -ge 8 -or -not [System.IO.File]::Exists($destinationExtended)){
      if([System.IO.File]::Exists($destinationExtended)){
        try{[System.IO.File]::Delete($destinationExtended)}catch{}
      }
      throw "Copy failed. Primary error: $firstError; Robocopy exit code: $exit"
    }
  }
}

$SamplePlan=[System.IO.Path]::GetFullPath($SamplePlan)
if(-not(Test-Path -LiteralPath $SamplePlan -PathType Leaf)){
  throw "Sample plan not found: $SamplePlan"
}

$SourceRoot=Normalize-Root $SourceRoot
$ExpectedSerial=$ExpectedSerial.Trim()
$DestinationRoot=Normalize-Root $DestinationRoot

$planDrive=[System.IO.Path]::GetPathRoot($SamplePlan)
$destDrive=[System.IO.Path]::GetPathRoot($DestinationRoot)

if($planDrive.TrimEnd('\').Equals($SourceRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: sample plan must not be stored on the source volume.'
}

if($destDrive.TrimEnd('\').Equals($SourceRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: destination must not be on the source volume.'
}

$disk=Get-CheckedDisk $SourceRoot $ExpectedSerial
$DiskNumber=[int]$disk.Number
$DiskFriendlyName=[string]$disk.FriendlyName
$DiskReadOnlyAtStart=[bool]$disk.IsReadOnly
$ExpectedSourceId='DISK-'+$ExpectedSerial

Write-Host
Write-Host 'RECOVERY AUDIT COPY - PREFLIGHT'
Write-Host '==============================='
Write-Host "Plan        : $SamplePlan"
Write-Host "Source      : $SourceRoot"
Write-Host "Disk        : $DiskFriendlyName"
Write-Host "Read-only   : $DiskReadOnlyAtStart"
Write-Host "Destination : $DestinationRoot"
Write-Host

$rows=Import-Csv -LiteralPath $SamplePlan -Delimiter ';' -Encoding UTF8
if($null -eq $rows -or $rows.Count -eq 0){
  throw 'Sample plan contains no rows.'
}

$required=@('PlanNo','FileID','SourceID','OriginalPath','RelativePath','SizeBytes','LastWriteTime','TypeGroup','TimeBand')
foreach($name in $required){
  if(-not ($rows[0].PSObject.Properties.Name -contains $name)){
    throw "Required plan column missing: $name"
  }
}

$validated=New-Object System.Collections.Generic.List[object]
$destSet=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$fileIdSet=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

[UInt64]$PlanFiles=0
[UInt64]$PlanBytes=0

foreach($r in $rows){
  if($r.SourceID -ne $ExpectedSourceId){
    throw "SAFETY ABORT: FileID $($r.FileID) has unexpected SourceID '$($r.SourceID)'."
  }

  if([string]::IsNullOrWhiteSpace($r.FileID)){
    throw 'Plan contains a blank FileID.'
  }

  if(-not $fileIdSet.Add($r.FileID)){
    throw "Plan contains duplicate FileID: $($r.FileID)"
  }

  $source=[System.IO.Path]::GetFullPath($r.OriginalPath)
  if(-not $source.StartsWith($SourceRoot,[System.StringComparison]::OrdinalIgnoreCase)){
    throw "SAFETY ABORT: source path is outside SourceRoot: $source"
  }

  [UInt64]$size=0
  if(-not [UInt64]::TryParse($r.SizeBytes,[ref]$size)){
    throw "Invalid SizeBytes for FileID $($r.FileID): $($r.SizeBytes)"
  }

  $stamp=[datetime]::MinValue
  if(-not [datetime]::TryParseExact(
    $r.LastWriteTime,
    'yyyy-MM-dd HH:mm:ss',
    [System.Globalization.CultureInfo]::InvariantCulture,
    [System.Globalization.DateTimeStyles]::None,
    [ref]$stamp
  )){
    throw "Invalid LastWriteTime for FileID $($r.FileID): $($r.LastWriteTime)"
  }

  $destination=Get-CanonicalDestination $DestinationRoot $r.RelativePath
  if(-not $destSet.Add($destination)){
    throw "Plan contains a destination collision: $destination"
  }

  $validated.Add([pscustomobject]@{
    PlanNo=$r.PlanNo
    FileID=$r.FileID
    SourceID=$r.SourceID
    SourcePath=$source
    RelativePath=$r.RelativePath
    DestinationPath=$destination
    SizeBytes=$size
    LastWriteTime=$stamp
    LastWriteTimeText=$r.LastWriteTime
    TypeGroup=$r.TypeGroup
    TimeBand=$r.TimeBand
  })

  $PlanFiles++
  $PlanBytes+=$size
}

if($ExpectedPlanFiles -gt 0 -and $PlanFiles -ne $ExpectedPlanFiles){
  throw "PRECHECK ABORT: plan has $PlanFiles files but ExpectedPlanFiles is $ExpectedPlanFiles."
}

if($ExpectedPlanBytes -gt 0 -and $PlanBytes -ne $ExpectedPlanBytes){
  throw "PRECHECK ABORT: plan has $PlanBytes bytes but ExpectedPlanBytes is $ExpectedPlanBytes."
}

if(Test-Path -LiteralPath $DestinationRoot){
  $existing=@(Get-ChildItem -LiteralPath $DestinationRoot -Force -ErrorAction Stop | Select-Object -First 1)
  if($existing.Count -gt 0){
    throw "SAFETY ABORT: destination already exists and is not empty: $DestinationRoot"
  }
}else{
  New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null
}

$destDriveName=$destDrive.Substring(0,1)
$drive=Get-PSDrive -Name $destDriveName -ErrorAction Stop
[UInt64]$requiredFree=$PlanBytes+$MinimumFreeReserveBytes

if([UInt64]$drive.Free -lt $requiredFree){
  throw ("PRECHECK ABORT: destination has {0:N0} free bytes; at least {1:N0} are required." -f [UInt64]$drive.Free,$requiredFree)
}

$ManifestPath=Join-Path $DestinationRoot 'copy-manifest.csv'
$SummaryPath=Join-Path $DestinationRoot 'copy-summary.txt'
$InProgressPath=Join-Path $DestinationRoot 'COPY_IN_PROGRESS.txt'
$PlanCopyPath=Join-Path $DestinationRoot 'approved-sample-plan.csv'

foreach($reserved in @($ManifestPath,$SummaryPath,$InProgressPath,$PlanCopyPath)){
  if($destSet.Contains($reserved)){
    throw "PRECHECK ABORT: planned sample collides with toolkit metadata path: $reserved"
  }
}

[System.IO.File]::Copy($SamplePlan,$PlanCopyPath,$false)

@(
  'RECOVERY AUDIT COPY IN PROGRESS'
  "Started=$($RunStarted.ToString('yyyy-MM-dd HH:mm:ss'))"
  "Plan=$SamplePlan"
  "PlanFiles=$PlanFiles"
  "PlanBytes=$PlanBytes"
) | Set-Content -LiteralPath $InProgressPath -Encoding UTF8

$utf8=New-Object System.Text.UTF8Encoding($true)
$manifest=New-Object System.IO.StreamWriter($ManifestPath,$false,$utf8,1048576)
$manifest.WriteLine('PlanNo;FileID;SourceID;SourcePath;RelativePath;DestinationPath;ExpectedSizeBytes;CopiedSizeBytes;LastWriteTime;TypeGroup;TimeBand;Result;Error')

[UInt64]$ProcessedFiles=0
[UInt64]$ProcessedBytes=0
[UInt64]$CopiedFiles=0
[UInt64]$CopiedBytes=0
[UInt64]$FailedFiles=0
[UInt64]$FailedBytes=0

$lastUi=[datetime]::MinValue
$lastSafety=[datetime]::MinValue
$copyStarted=Get-Date
$aborted=$false
$abortMessage=''

Write-Host ("Plan files  : {0:N0}" -f $PlanFiles)
Write-Host ("Plan bytes  : {0:N0} ({1:N3} decimal GB)" -f $PlanBytes,(Format-DecimalGB $PlanBytes))
Write-Host ("Free bytes  : {0:N0}" -f [UInt64]$drive.Free)
Write-Host
Write-Host 'Copying approved sample ...'

try{
  foreach($r in $validated){
    $now=Get-Date

    if(($now-$lastSafety).TotalSeconds -ge 30){
      [void](Assert-ReadOnly $DiskNumber $ExpectedSerial)
      $lastSafety=$now
    }

    [UInt64]$copiedSize=0
    $result='COPIED'
    $errorText=''

    try{
      Copy-FileNoOverwrite $r.SourcePath $r.DestinationPath

      $copiedSize=Get-FileLengthExtended $r.DestinationPath

      if($copiedSize -ne [UInt64]$r.SizeBytes){
        throw "SIZE_MISMATCH expected=$($r.SizeBytes) actual=$copiedSize"
      }

      $CopiedFiles++
      $CopiedBytes+=$copiedSize
    }
    catch{
      $result='FAILED'
      $errorText=$_.Exception.Message
      $FailedFiles++
      $FailedBytes+=[UInt64]$r.SizeBytes

      try{Remove-FileExtended $r.DestinationPath}catch{}
    }

    $ProcessedFiles++
    $ProcessedBytes+=[UInt64]$r.SizeBytes

    $manifest.WriteLine((@(
      (Csv $r.PlanNo),
      (Csv $r.FileID),
      (Csv $r.SourceID),
      (Csv $r.SourcePath),
      (Csv $r.RelativePath),
      (Csv $r.DestinationPath),
      ([string]$r.SizeBytes),
      ([string]$copiedSize),
      (Csv $r.LastWriteTimeText),
      (Csv $r.TypeGroup),
      (Csv $r.TimeBand),
      (Csv $result),
      (Csv $errorText)
    ) -join ';'))

    if(($ProcessedFiles % 100) -eq 0){
      $manifest.Flush()
    }

    $now=Get-Date
    if(($now-$lastUi).TotalMilliseconds -ge 250 -or $ProcessedFiles -eq $PlanFiles){
      $elapsed=$now-$copyStarted
      $percent=if($PlanBytes -gt 0){([double]$ProcessedBytes/[double]$PlanBytes)*100}else{100}
      $rate=if($elapsed.TotalSeconds -gt 0){[double]$ProcessedBytes/$elapsed.TotalSeconds}else{0}
      $eta=[TimeSpan]::Zero

      if($rate -gt 0 -and $ProcessedBytes -lt $PlanBytes){
        $eta=[TimeSpan]::FromSeconds(([double]($PlanBytes-$ProcessedBytes))/$rate)
      }

      $show=$r.SourcePath
      if($show.Length -gt 150){$show='...'+$show.Substring($show.Length-147)}

      $status=('{0:N0}/{1:N0} files | {2:N3}/{3:N3} GB | {4:N1}% | {5:N1} MB/s | ETA {6} | failed {7:N0}' -f $ProcessedFiles,$PlanFiles,([double]$ProcessedBytes/1000000000),([double]$PlanBytes/1000000000),$percent,($rate/1000000),(Format-Duration $eta),$FailedFiles)

      $progressPercent=[math]::Max(0,[math]::Min(100,$percent))
      Write-Progress -Activity 'Recovery Audit Toolkit - controlled copy' -Status $status -CurrentOperation $show -PercentComplete $progressPercent

      $lastUi=$now
    }
  }
}
catch{
  $aborted=$true
  $abortMessage=$_.Exception.Message
}
finally{
  $manifest.Flush()
  $manifest.Dispose()
  Write-Progress -Activity 'Recovery Audit Toolkit - controlled copy' -Completed
}

$DiskReadOnlyAtEnd=$false
try{
  $diskEnd=Assert-ReadOnly $DiskNumber $ExpectedSerial
  $DiskReadOnlyAtEnd=[bool]$diskEnd.IsReadOnly
}
catch{
  if(-not $aborted){
    $aborted=$true
    $abortMessage=$_.Exception.Message
  }
}

$finished=Get-Date
$duration=$finished-$RunStarted
$status=if($aborted){'ABORTED'}elseif($FailedFiles -gt 0){'COMPLETE_WITH_ERRORS'}else{'COMPLETE'}

$summary=@(
  'RECOVERY AUDIT COPY'
  '==================='
  "Version=$ScriptVersion"
  "Status=$status"
  "Started=$($RunStarted.ToString('yyyy-MM-dd HH:mm:ss'))"
  "Finished=$($finished.ToString('yyyy-MM-dd HH:mm:ss'))"
  "Duration=$(Format-Duration $duration)"
  "SamplePlan=$SamplePlan"
  "SourceRoot=$SourceRoot"
  "DiskFriendlyName=$DiskFriendlyName"
  "DiskReadOnlyAtStart=$DiskReadOnlyAtStart"
  "DiskReadOnlyAtEnd=$DiskReadOnlyAtEnd"
  "DestinationRoot=$DestinationRoot"
  "PlanFiles=$PlanFiles"
  "PlanBytes=$PlanBytes"
  "ProcessedFiles=$ProcessedFiles"
  "ProcessedBytes=$ProcessedBytes"
  "CopiedFiles=$CopiedFiles"
  "CopiedBytes=$CopiedBytes"
  "FailedFiles=$FailedFiles"
  "FailedBytes=$FailedBytes"
  "Manifest=$ManifestPath"
  "ApprovedPlanCopy=$PlanCopyPath"
  "SourceHashesCalculated=False"
  "OverwriteAllowed=False"
  "AbortMessage=$abortMessage"
)

$summary | Set-Content -LiteralPath $SummaryPath -Encoding UTF8

if(-not $aborted){
  Remove-Item -LiteralPath $InProgressPath -Force -ErrorAction SilentlyContinue
}

Write-Host
Write-Host 'RECOVERY AUDIT COPY - DONE'
Write-Host '=========================='
Write-Host "Status                  : $status"
Write-Host ("Processed files         : {0:N0} / {1:N0}" -f $ProcessedFiles,$PlanFiles)
Write-Host ("Copied files            : {0:N0}" -f $CopiedFiles)
Write-Host ("Copied decimal GB       : {0:N3}" -f ([double]$CopiedBytes/1000000000))
Write-Host ("Failed files            : {0:N0}" -f $FailedFiles)
Write-Host ("Source read-only        : {0}" -f $DiskReadOnlyAtEnd)
Write-Host "Manifest                : $ManifestPath"
Write-Host "Summary                 : $SummaryPath"

if($aborted){
  throw "CONTROLLED COPY ABORTED: $abortMessage"
}
