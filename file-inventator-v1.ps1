#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
  Read-only inventory of a mounted Windows volume.

.DESCRIPTION
  file-inventator-v1.ps1 inventories all regular files that Robocopy can
  enumerate with backup rights. It is intended for a source disk that Windows
  already reports as read-only.

  V1:
    - uses Robocopy /L, so it lists only and never copies source files;
    - uses /B to enumerate ACL-protected directories without taking ownership;
    - excludes junction targets (/XJ) and symbolic-link targets (/SL);
    - calculates no source-file hashes;
    - writes all results to a timestamped folder on the current user's Desktop;
    - performs a pre-count, a full Unicode listing, then CSV validation.

  A successful run ends with:
    Status=COMPLETE
    Validation=PASS

.SAFETY
  The script aborts unless the supplied source serial number matches and the
  source disk reports IsReadOnly=True. The Desktop must be on another drive.

  Software read-only is not a substitute for a hardware write blocker when
  strict forensic preservation is required.

.SCOPE
  V1 does not recover deleted files, inspect unallocated space, carve files,
  inspect other unmounted/RAW partitions, or calculate content hashes.

.EXAMPLE
  .\file-inventator-v1.ps1 -SourceRoot 'X:\' -ExpectedSerial 'SERIAL_NUMBER'
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidatePattern('^[A-Za-z]:\\?$')]
  [string]$SourceRoot,

  [Parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [string]$ExpectedSerial
)

$ErrorActionPreference = 'Stop'
$ScriptName = 'file-inventator-v1'
$ScriptVersion = '1.0.2'
$RunStarted = Get-Date
$RunID = $RunStarted.ToString('yyyyMMdd_HHmmss')

# Locale-independent file-line parser:
# <status text> <bytes> yyyy/MM/dd HH:mm:ss <absolute path>
$FileLineRegex = [regex]::new(
  '^\s*(?<Class>.*?)\s+(?<Size>\d+)\s+(?<Date>\d{4}/\d{2}/\d{2})\s+(?<Time>\d{2}:\d{2}:\d{2})\s+(?<Path>(?:\\\\\?\\)?[A-Za-z]:\\.*)

function Normalize-Root([string]$Path) {
  $p = [System.IO.Path]::GetFullPath($Path)
  if (-not $p.EndsWith('\')) { $p += '\' }
  if ($p -notmatch '^[A-Za-z]:\\$') { throw "SourceRoot must be a drive root such as X:\. Received: $Path" }
  return $p
}

function Format-Duration([TimeSpan]$Span) {
  $h = [int][math]::Floor($Span.TotalHours)
  return ('{0:00}:{1:00}:{2:00}' -f $h,$Span.Minutes,$Span.Seconds)
}

function Csv([AllowNull()][string]$Value) {
  if ($null -eq $Value) { return '""' }
  return '"' + $Value.Replace('"','""') + '"'
}

function Get-CheckedDisk([string]$Root,[string]$Serial) {
  $letter = $Root.Substring(0,1)
  $partition = Get-Partition -DriveLetter $letter -ErrorAction Stop
  $disk = $partition | Get-Disk -ErrorAction Stop
  $actual = ([string]$disk.SerialNumber).Trim()

  if ($actual -ne $Serial.Trim()) {
    throw "SAFETY ABORT: source serial '$actual' does not match expected serial '$Serial'."
  }
  if (-not $disk.IsReadOnly) {
    throw "SAFETY ABORT: source disk $($disk.Number) is NOT read-only."
  }
  return $disk
}

function Assert-ReadOnly([int]$DiskNumber,[string]$Serial) {
  $disk = Get-Disk -Number $DiskNumber -ErrorAction Stop
  if (([string]$disk.SerialNumber).Trim() -ne $Serial.Trim()) {
    throw 'SAFETY ABORT: source disk identity changed during the run.'
  }
  if (-not $disk.IsReadOnly) {
    throw 'SAFETY ABORT: source disk is no longer read-only.'
  }
  return $disk
}

function Quote-Arg([string]$Value) {
  if ($Value -match '\s') { return '"' + $Value + '"' }
  return $Value
}

function Start-Robo([string[]]$RoboArguments) {
  $exe = Join-Path $env:SystemRoot 'System32\robocopy.exe'
  if (-not (Test-Path -LiteralPath $exe)) { throw "Robocopy not found: $exe" }

  if ($null -eq $RoboArguments -or $RoboArguments.Count -eq 0) {
    throw 'Internal error: no Robocopy arguments were supplied.'
  }

  $line = (($RoboArguments | ForEach-Object { Quote-Arg $_ }) -join ' ')

  if ([string]::IsNullOrWhiteSpace($line)) {
    throw 'Internal error: Robocopy argument line is empty.'
  }

  return Start-Process -FilePath $exe -ArgumentList $line -NoNewWindow -PassThru
}

function Wait-Robo([System.Diagnostics.Process]$Process,[string]$Activity,[int]$DiskNumber,[string]$Serial) {
  $start = Get-Date
  $lastCheck = [datetime]::MinValue

  while (-not $Process.HasExited) {
    $now = Get-Date
    Write-Progress -Activity $Activity -Status ("Elapsed {0}" -f (Format-Duration ($now-$start))) -CurrentOperation $script:SourceRoot

    if (($now-$lastCheck).TotalSeconds -ge 30) {
      try { [void](Assert-ReadOnly $DiskNumber $Serial) }
      catch { try { $Process.Kill() } catch {}; throw }
      $lastCheck = $now
    }

    Start-Sleep -Milliseconds 750
    $Process.Refresh()
  }

  $Process.WaitForExit()
  Write-Progress -Activity $Activity -Completed
}

function Get-RoboSummary([string]$LogPath) {
  $tail = Get-Content -LiteralPath $LogPath -Encoding Unicode -Tail 160
  $rows = New-Object System.Collections.Generic.List[object]

  foreach ($line in $tail) {
    if ($line -match '^\s*[^:]+:\s+(?<Total>\d+)\s+(?<Copied>\d+)\s+(?<Skipped>\d+)\s+(?<Mismatch>\d+)\s+(?<Failed>\d+)\s+(?<Extras>\d+)\s*$') {
      $rows.Add([pscustomobject]@{
        Total=[UInt64]$matches.Total
        Skipped=[UInt64]$matches.Skipped
        Failed=[UInt64]$matches.Failed
      })
    }
  }

  if ($rows.Count -lt 3) { throw "Could not parse Robocopy summary from $LogPath" }

  return [pscustomobject]@{
    DirsTotal=$rows[0].Total
    DirsSkipped=$rows[0].Skipped
    DirsFailed=$rows[0].Failed
    FilesTotal=$rows[1].Total
    FilesSkipped=$rows[1].Skipped
    FilesFailed=$rows[1].Failed
    BytesTotal=$rows[2].Total
    BytesSkipped=$rows[2].Skipped
    BytesFailed=$rows[2].Failed
  }
}

function Watch-Raw(
  [System.Diagnostics.Process]$Process,
  [string]$LogPath,
  [UInt64]$ExpectedFiles,
  [UInt64]$ExpectedBytes,
  [int]$DiskNumber,
  [string]$Serial
) {
  while (-not (Test-Path -LiteralPath $LogPath)) {
    if ($Process.HasExited) { throw 'Robocopy exited before creating the raw log.' }
    Start-Sleep -Milliseconds 250
  }

  $stream = New-Object System.IO.FileStream($LogPath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
  $reader = New-Object System.IO.StreamReader($stream,[System.Text.Encoding]::Unicode,$true,65536)

  [UInt64]$seen=0
  [UInt64]$bytes=0
  [UInt64]$nextMark=100000
  $start=Get-Date
  $lastUi=[datetime]::MinValue
  $lastCheck=[datetime]::MinValue
  $current=''

  try {
    while ($true) {
      $line=$reader.ReadLine()

      if ($null -ne $line) {
        $m=$script:FileLineRegex.Match($line)

        if ($m.Success) {
          $seen++
          $bytes += [UInt64]$m.Groups['Size'].Value
          $current=$m.Groups['Path'].Value
          $now=Get-Date

          if (($now-$lastUi).TotalMilliseconds -ge 500) {
            $elapsed=$now-$start
            $pf=if($ExpectedFiles -gt 0){([double]$seen/[double]$ExpectedFiles)*100}else{0}
            $pb=if($ExpectedBytes -gt 0){([double]$bytes/[double]$ExpectedBytes)*100}else{0}
            $rate=if($elapsed.TotalSeconds -gt 0){[double]$seen/$elapsed.TotalSeconds}else{0}
            $eta=[TimeSpan]::Zero

            if($rate -gt 0 -and $seen -lt $ExpectedFiles){
              $eta=[TimeSpan]::FromSeconds(([double]($ExpectedFiles-$seen))/$rate)
            }

            $show=$current
            if($show.Length -gt 160){$show='...'+$show.Substring($show.Length-157)}

            $status='{0:N0}/{1:N0} files | files {2:N1}% | bytes {3:N1}% | {4:N1} files/s | ETA {5}' -f $seen,$ExpectedFiles,$pf,$pb,$rate,(Format-Duration $eta)
            Write-Progress -Activity 'Phase 2/3 - Raw listing' -Status $status -CurrentOperation $show -PercentComplete ([math]::Max(0,[math]::Min(100,$pf)))
            $lastUi=$now
          }

          if($seen -ge $nextMark){
            $msg='{0} | {1:N0}/{2:N0} files | {3:N1}% | {4:N2} GiB' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),$seen,$ExpectedFiles,(([double]$seen/[double]$ExpectedFiles)*100),([double]$bytes/1GB)
            Write-Host $msg
            Add-Content -LiteralPath $script:ProgressLog -Value $msg -Encoding UTF8
            while($nextMark -le $seen){$nextMark+=100000}
          }
        }
        continue
      }

      if($Process.HasExited){break}

      $now=Get-Date
      if(($now-$lastCheck).TotalSeconds -ge 30){
        try{[void](Assert-ReadOnly $DiskNumber $Serial)}
        catch{try{$Process.Kill()}catch{};throw}
        $lastCheck=$now
      }

      Start-Sleep -Milliseconds 200
      $Process.Refresh()
    }
  }
  finally {
    $reader.Dispose()
    $stream.Dispose()
    Write-Progress -Activity 'Phase 2/3 - Raw listing' -Completed
  }

  return [pscustomobject]@{FilesObserved=$seen;BytesObserved=$bytes}
}

function Parse-RawToCsv(
  [string]$RawLog,
  [string]$CsvPath,
  [string]$ParserErrors,
  [string]$Root,
  [string]$SourceId,
  [UInt64]$ExpectedFiles,
  [UInt64]$ExpectedBytes
) {
  $utf8bom=New-Object System.Text.UTF8Encoding($true)
  $reader=New-Object System.IO.StreamReader($RawLog,[System.Text.Encoding]::Unicode,$true,1048576)
  $writer=New-Object System.IO.StreamWriter($CsvPath,$false,$utf8bom,1048576)
  $errWriter=New-Object System.IO.StreamWriter($ParserErrors,$false,$utf8bom,65536)

  [UInt64]$n=0
  [UInt64]$sum=0
  [UInt64]$parseErr=0
  [UInt64]$nextMark=100000
  $start=Get-Date
  $lastUi=[datetime]::MinValue

  try {
    $writer.WriteLine('FileID;SourceID;RobocopyClass;FullPath;RelativePath;RelativeDirectory;FileName;BaseName;Extension;SizeBytes;LastWriteTime;LastWriteYear;TopLevelDirectory;PathLength;IsZeroByte')

    while(($line=$reader.ReadLine()) -ne $null){
      $m=$script:FileLineRegex.Match($line)

      if($m.Success){
        $n++
        $sizeText=$m.Groups['Size'].Value
        $size=[UInt64]$sizeText
        $sum+=$size

        $roboClass=$m.Groups['Class'].Value.Trim()
        $full=$m.Groups['Path'].Value
        $path=$full
        if($path.StartsWith('\\?\')){$path=$path.Substring(4)}

        if(-not $path.StartsWith($Root,[System.StringComparison]::OrdinalIgnoreCase)){
          $parseErr++
          $errWriter.WriteLine("SOURCE_MISMATCH|$line")
          continue
        }

        $rel=$path.Substring($Root.Length)
        $slash=$path.LastIndexOf('\')

        if($slash -ge 0){
          $dir=$path.Substring(0,$slash)
          $name=$path.Substring($slash+1)
        }else{
          $dir=''
          $name=$path
        }

        if($dir.Length -ge $Root.Length){$relDir=$dir.Substring($Root.Length).TrimStart('\')}else{$relDir=''}

        $first=$rel.IndexOf('\')
        if($first -gt 0){$top=$rel.Substring(0,$first)}else{$top='[ROOT]'}

        $dot=$name.LastIndexOf('.')
        if($dot -gt 0 -and $dot -lt ($name.Length-1)){
          $base=$name.Substring(0,$dot)
          $ext=$name.Substring($dot).ToLowerInvariant()
        }else{
          $base=$name
          $ext=''
        }

        $date=$m.Groups['Date'].Value
        $time=$m.Groups['Time'].Value
        $stamp=$date.Replace('/','-')+' '+$time
        $year=$date.Substring(0,4)
        $id='F'+$n.ToString('D9')
        $zero=if($size -eq 0){'1'}else{'0'}

        $fields=@(
          (Csv $id),(Csv $SourceId),(Csv $roboClass),(Csv $path),(Csv $rel),(Csv $relDir),
          (Csv $name),(Csv $base),(Csv $ext),$sizeText,(Csv $stamp),$year,
          (Csv $top),$path.Length,$zero
        )

        $writer.WriteLine(($fields -join ';'))

        $now=Get-Date
        if(($now-$lastUi).TotalMilliseconds -ge 500){
          $elapsed=$now-$start
          $pf=if($ExpectedFiles -gt 0){([double]$n/[double]$ExpectedFiles)*100}else{0}
          $pb=if($ExpectedBytes -gt 0){([double]$sum/[double]$ExpectedBytes)*100}else{0}
          $rate=if($elapsed.TotalSeconds -gt 0){[double]$n/$elapsed.TotalSeconds}else{0}
          $eta=[TimeSpan]::Zero
          if($rate -gt 0 -and $n -lt $ExpectedFiles){$eta=[TimeSpan]::FromSeconds(([double]($ExpectedFiles-$n))/$rate)}
          $status='{0:N0}/{1:N0} files | {2:N1}% | bytes {3:N1}% | ETA {4}' -f $n,$ExpectedFiles,$pf,$pb,(Format-Duration $eta)
          Write-Progress -Activity 'Phase 3/3 - Master CSV' -Status $status -CurrentOperation $path -PercentComplete ([math]::Max(0,[math]::Min(100,$pf)))
          $lastUi=$now
        }

        if($n -ge $nextMark){
          $msg='{0} | CSV {1:N0}/{2:N0} files | {3:N1}% | {4:N2} GiB' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),$n,$ExpectedFiles,(([double]$n/[double]$ExpectedFiles)*100),([double]$sum/1GB)
          Write-Host $msg
          Add-Content -LiteralPath $script:ProgressLog -Value $msg -Encoding UTF8
          while($nextMark -le $n){$nextMark+=100000}
        }
      }
      elseif($line -match [regex]::Escape($Root) -and $line -match '\d{4}/\d{2}/\d{2}' -and $line -match '\d{2}:\d{2}:\d{2}'){
        $parseErr++
        $errWriter.WriteLine("UNPARSED|$line")
      }
    }
  }
  finally {
    $writer.Flush();$writer.Dispose()
    $errWriter.Flush();$errWriter.Dispose()
    $reader.Dispose()
    Write-Progress -Activity 'Phase 3/3 - Master CSV' -Completed
  }

  return [pscustomobject]@{Records=$n;Bytes=$sum;ParserErrors=$parseErr}
}

function Export-RoboErrors([string[]]$Logs,[string]$Output) {
  $utf8bom=New-Object System.Text.UTF8Encoding($true)
  $writer=New-Object System.IO.StreamWriter($Output,$false,$utf8bom,65536)
  [UInt64]$count=0

  try{
    foreach($log in $Logs){
      if(-not(Test-Path -LiteralPath $log)){continue}
      $reader=New-Object System.IO.StreamReader($log,[System.Text.Encoding]::Unicode,$true,65536)
      try{
        while(($line=$reader.ReadLine()) -ne $null){
          if($line -match '(?i)^\s*(?:(?:\d{4}/\d{2}/\d{2})\s+\d{2}:\d{2}:\d{2}\s+)?(?:FEHLER|ERROR)(?:\s+\d+\s+\(0x[0-9A-F]+\)|\s*:)' ){
            $count++
            $writer.WriteLine("[$([System.IO.Path]::GetFileName($log))] $line")
          }
        }
      }finally{$reader.Dispose()}
    }
  }finally{$writer.Flush();$writer.Dispose()}

  return $count
}

function Write-Manifest([string]$Status,[string]$Validation,[hashtable]$Extra=@{}) {
  $lines=New-Object System.Collections.Generic.List[string]
  $lines.Add('FILE INVENTATOR V1 - MANIFEST')
  $lines.Add('=============================')
  $lines.Add('')
  $lines.Add("ScriptVersion=$script:ScriptVersion")
  $lines.Add("RunID=$script:RunID")
  $lines.Add("Status=$Status")
  $lines.Add("Validation=$Validation")
  $lines.Add("Started=$($script:RunStarted.ToString('yyyy-MM-dd HH:mm:ss'))")
  $lines.Add("Updated=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
  $lines.Add('')
  $lines.Add("SourceRoot=$script:SourceRoot")
  $lines.Add("SourceID=$script:SourceID")
  $lines.Add("DiskNumber=$script:DiskNumber")
  $lines.Add("DiskFriendlyName=$script:DiskFriendlyName")
  $lines.Add("DiskSerialNumber=$script:DiskSerialNumber")
  $lines.Add("DiskUniqueId=$script:DiskUniqueId")
  $lines.Add("DiskPartitionStyle=$script:DiskPartitionStyle")
  $lines.Add("DiskSizeBytes=$script:DiskSizeBytes")
  $lines.Add("DiskReadOnlyAtStart=$script:DiskReadOnlyAtStart")
  $lines.Add("VolumeLabel=$script:VolumeLabel")
  $lines.Add("VolumeFileSystem=$script:VolumeFileSystem")
  $lines.Add("VolumeSizeBytes=$script:VolumeSizeBytes")
  $lines.Add('')
  $lines.Add('Enumeration=ROBOCOPY /L /B /E /XJ /SL; full listing additionally uses /V')
  $lines.Add('VerboseSkippedFiles=True')
  $lines.Add('ContentHashesCalculated=False')
  $lines.Add('JunctionTargetsFollowed=False')
  $lines.Add('SymbolicLinkTargetsFollowed=False')
  $lines.Add('DeletedFilesRecovered=False')
  $lines.Add('UnallocatedSpaceScanned=False')
  $lines.Add('')
  $lines.Add("OutputDirectory=$script:OutputDir")
  $lines.Add("MasterCsv=$script:MasterCsv")
  $lines.Add("RawListing=$script:RawLog")
  $lines.Add("PrecountLog=$script:PrecountLog")
  $lines.Add("ErrorsLog=$script:ErrorsLog")
  $lines.Add("ParserErrorsLog=$script:ParserErrorsLog")

  foreach($key in ($Extra.Keys|Sort-Object)){$lines.Add("$key=$($Extra[$key])")}
  $lines|Set-Content -LiteralPath $script:ManifestPath -Encoding UTF8
}

# Initial safety checks
$SourceRoot=Normalize-Root $SourceRoot
$ExpectedSerial=$ExpectedSerial.Trim()
$disk=Get-CheckedDisk $SourceRoot $ExpectedSerial
$letter=$SourceRoot.Substring(0,1)
$volume=Get-Volume -DriveLetter $letter -ErrorAction Stop

$DiskNumber=[int]$disk.Number
$DiskFriendlyName=[string]$disk.FriendlyName
$DiskSerialNumber=([string]$disk.SerialNumber).Trim()
$DiskUniqueId=[string]$disk.UniqueId
$DiskPartitionStyle=[string]$disk.PartitionStyle
$DiskSizeBytes=[UInt64]$disk.Size
$DiskReadOnlyAtStart=[bool]$disk.IsReadOnly
$VolumeLabel=[string]$volume.FileSystemLabel
$VolumeFileSystem=[string]$volume.FileSystem
$VolumeSizeBytes=[UInt64]$volume.Size
$SourceID='DISK-'+$DiskSerialNumber

$Desktop=[Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
if([string]::IsNullOrWhiteSpace($Desktop)){throw 'Could not resolve Desktop path.'}

$desktopRoot=[System.IO.Path]::GetPathRoot($Desktop)
if($desktopRoot.TrimEnd('\').Equals($SourceRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: Desktop output and source volume are on the same drive.'
}

$OutputDir=Join-Path $Desktop ($ScriptName+'_'+$RunID)
$i=2
while(Test-Path -LiteralPath $OutputDir){
  $OutputDir=Join-Path $Desktop (($ScriptName+'_'+$RunID)+'_{0:D2}' -f $i)
  $i++
}
New-Item -ItemType Directory -Path $OutputDir|Out-Null

$DummyTarget=Join-Path $OutputDir '_robocopy_dummy_target'
$PrecountLog=Join-Path $OutputDir ($ScriptName+'_precount.txt')
$RawLog=Join-Path $OutputDir ($ScriptName+'_raw.txt')
$MasterCsv=Join-Path $OutputDir ($ScriptName+'_master.csv')
$ErrorsLog=Join-Path $OutputDir ($ScriptName+'_errors.txt')
$ParserErrorsLog=Join-Path $OutputDir ($ScriptName+'_parser-errors.txt')
$ProgressLog=Join-Path $OutputDir ($ScriptName+'_progress.txt')
$ManifestPath=Join-Path $OutputDir ($ScriptName+'_manifest.txt')

New-Item -ItemType Directory -Path $DummyTarget|Out-Null
New-Item -ItemType File -Path $ProgressLog|Out-Null
Write-Manifest 'IN_PROGRESS' 'PENDING'

Write-Host
Write-Host 'FILE INVENTATOR V1'
Write-Host '=================='
Write-Host "Source       : $SourceRoot"
Write-Host "Disk         : $DiskFriendlyName"
Write-Host "Serial       : $DiskSerialNumber"
Write-Host "Read-only    : $DiskReadOnlyAtStart"
Write-Host "Output       : $OutputDir"
Write-Host "Content hash : NO"
Write-Host

try{
  # Phase 1
  Write-Host 'Phase 1/3: Pre-count with backup rights ...'

  $preArgs=@($SourceRoot,$DummyTarget,'/L','/B','/E','/XJ','/SL','/R:0','/W:0','/BYTES','/NFL','/NDL','/NP',"/UNILOG:$PrecountLog")
  $pre=Start-Robo $preArgs
  Wait-Robo $pre 'Phase 1/3 - Pre-count' $DiskNumber $ExpectedSerial

  $preExit=[int]$pre.ExitCode
  $preSum=Get-RoboSummary $PrecountLog

  if($preExit -ge 8){throw "Robocopy precount exit code $preExit"}
  if($preSum.FilesFailed -gt 0 -or $preSum.DirsFailed -gt 0){throw 'Robocopy precount reported failures.'}

  [UInt64]$ExpectedFiles=$preSum.FilesTotal
  [UInt64]$ExpectedBytes=$preSum.BytesTotal
  if($ExpectedFiles -eq 0){throw 'No files found.'}

  Write-Host ("Expected files : {0:N0}" -f $ExpectedFiles)
  Write-Host ("Expected bytes : {0:N0}" -f $ExpectedBytes)
  Write-Host ("Expected GiB   : {0:N3}" -f ([double]$ExpectedBytes/1GB))
  Write-Host

  $estimate=([double]$ExpectedFiles*1500.0)+200MB
  $outDrive=New-Object System.IO.DriveInfo($desktopRoot)
  if([double]$outDrive.AvailableFreeSpace -lt $estimate){throw 'Insufficient free space at output location.'}

  [void](Assert-ReadOnly $DiskNumber $ExpectedSerial)

  # Phase 2
  Write-Host 'Phase 2/3: Creating Unicode raw listing ...'

  $fullArgs=@($SourceRoot,$DummyTarget,'/L','/B','/E','/XJ','/SL','/V','/R:0','/W:0','/BYTES','/FP','/TS','/NP','/NDL',"/UNILOG:$RawLog")
  $full=Start-Robo $fullArgs
  $live=Watch-Raw $full $RawLog $ExpectedFiles $ExpectedBytes $DiskNumber $ExpectedSerial

  $full.WaitForExit()
  $fullExit=[int]$full.ExitCode
  $fullSum=Get-RoboSummary $RawLog

  if($fullExit -ge 8){throw "Robocopy full-run exit code $fullExit"}
  if($fullSum.FilesFailed -gt 0 -or $fullSum.DirsFailed -gt 0){throw 'Robocopy full run reported failures.'}

  Write-Host ("Raw listing completed: {0:N0} files." -f $fullSum.FilesTotal)

  # Phase 3
  Write-Host 'Phase 3/3: Building master CSV and validating ...'

  $parsed=Parse-RawToCsv $RawLog $MasterCsv $ParserErrorsLog $SourceRoot $SourceID $ExpectedFiles $ExpectedBytes
  $roboErrors=Export-RoboErrors @($PrecountLog,$RawLog) $ErrorsLog
  $diskEnd=Assert-ReadOnly $DiskNumber $ExpectedSerial

  $checks=[ordered]@{
    PrecountVsFullFiles=($preSum.FilesTotal -eq $fullSum.FilesTotal)
    PrecountVsFullBytes=($preSum.BytesTotal -eq $fullSum.BytesTotal)
    FullVsCsvFiles=($fullSum.FilesTotal -eq $parsed.Records)
    FullVsCsvBytes=($fullSum.BytesTotal -eq $parsed.Bytes)
    PrecountFailures0=(($preSum.FilesFailed+$preSum.DirsFailed) -eq 0)
    FullFailures0=(($fullSum.FilesFailed+$fullSum.DirsFailed) -eq 0)
    ParserErrors0=($parsed.ParserErrors -eq 0)
    PrecountExitOK=($preExit -lt 8)
    FullExitOK=($fullExit -lt 8)
    SourceStillReadOnly=([bool]$diskEnd.IsReadOnly)
  }

  $pass=$true
  foreach($v in $checks.Values){if(-not $v){$pass=$false;break}}

  if($pass){$status='COMPLETE';$validation='PASS'}
  else{$status='COMPLETE_WITH_VALIDATION_FAILURE';$validation='FAIL'}

  Write-Manifest $status $validation @{
    PrecountExitCode=$preExit
    FullRunExitCode=$fullExit
    PrecountDirs=$preSum.DirsTotal
    PrecountFiles=$preSum.FilesTotal
    PrecountBytes=$preSum.BytesTotal
    FullDirs=$fullSum.DirsTotal
    FullFiles=$fullSum.FilesTotal
    FullBytes=$fullSum.BytesTotal
    LiveFilesObserved=$live.FilesObserved
    LiveBytesObserved=$live.BytesObserved
    CsvRecords=$parsed.Records
    CsvBytes=$parsed.Bytes
    ParserErrors=$parsed.ParserErrors
    RobocopyErrorLines=$roboErrors
    Check_PrecountVsFullFiles=$checks.PrecountVsFullFiles
    Check_PrecountVsFullBytes=$checks.PrecountVsFullBytes
    Check_FullVsCsvFiles=$checks.FullVsCsvFiles
    Check_FullVsCsvBytes=$checks.FullVsCsvBytes
    Check_SourceStillReadOnly=$checks.SourceStillReadOnly
    ExcelRowLimitExceeded=($parsed.Records -gt 1048575)
    Finished=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  }

  try{Remove-Item -LiteralPath $DummyTarget -Force -Recurse -ErrorAction SilentlyContinue}catch{}

  Write-Host
  Write-Host 'DONE'
  Write-Host "Status      : $status"
  Write-Host "Validation  : $validation"
  Write-Host ("Files       : {0:N0}" -f $parsed.Records)
  Write-Host ("GiB         : {0:N3}" -f ([double]$parsed.Bytes/1GB))
  Write-Host ("Parser errs : {0:N0}" -f $parsed.ParserErrors)
  Write-Host ("Robo errors : {0:N0}" -f $roboErrors)
  Write-Host "Results     : $OutputDir"

  if($parsed.Records -gt 1048575){Write-Warning 'CSV exceeds one Excel worksheet row limit; the CSV itself remains complete.'}
  if($validation -ne 'PASS'){Write-Warning 'Completeness validation did not fully pass. Review manifest and error logs.'}
}
catch{
  $msg=$_.Exception.Message

  try{
    Write-Manifest 'FAILED' 'FAIL' @{
      FailureMessage=$msg
      FailedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
  }catch{}

  Write-Error "file-inventator-v1 aborted: $msg"
  Write-Host "Partial output remains at: $OutputDir"
  throw
}
,
  [System.Text.RegularExpressions.RegexOptions]::Compiled
)

function Normalize-Root([string]$Path) {
  $p = [System.IO.Path]::GetFullPath($Path)
  if (-not $p.EndsWith('\')) { $p += '\' }
  if ($p -notmatch '^[A-Za-z]:\\$') { throw "SourceRoot must be a drive root such as X:\. Received: $Path" }
  return $p
}

function Format-Duration([TimeSpan]$Span) {
  $h = [int][math]::Floor($Span.TotalHours)
  return ('{0:00}:{1:00}:{2:00}' -f $h,$Span.Minutes,$Span.Seconds)
}

function Csv([AllowNull()][string]$Value) {
  if ($null -eq $Value) { return '""' }
  return '"' + $Value.Replace('"','""') + '"'
}

function Get-CheckedDisk([string]$Root,[string]$Serial) {
  $letter = $Root.Substring(0,1)
  $partition = Get-Partition -DriveLetter $letter -ErrorAction Stop
  $disk = $partition | Get-Disk -ErrorAction Stop
  $actual = ([string]$disk.SerialNumber).Trim()

  if ($actual -ne $Serial.Trim()) {
    throw "SAFETY ABORT: source serial '$actual' does not match expected serial '$Serial'."
  }
  if (-not $disk.IsReadOnly) {
    throw "SAFETY ABORT: source disk $($disk.Number) is NOT read-only."
  }
  return $disk
}

function Assert-ReadOnly([int]$DiskNumber,[string]$Serial) {
  $disk = Get-Disk -Number $DiskNumber -ErrorAction Stop
  if (([string]$disk.SerialNumber).Trim() -ne $Serial.Trim()) {
    throw 'SAFETY ABORT: source disk identity changed during the run.'
  }
  if (-not $disk.IsReadOnly) {
    throw 'SAFETY ABORT: source disk is no longer read-only.'
  }
  return $disk
}

function Quote-Arg([string]$Value) {
  if ($Value -match '\s') { return '"' + $Value + '"' }
  return $Value
}

function Start-Robo([string[]]$RoboArguments) {
  $exe = Join-Path $env:SystemRoot 'System32\robocopy.exe'
  if (-not (Test-Path -LiteralPath $exe)) { throw "Robocopy not found: $exe" }

  if ($null -eq $RoboArguments -or $RoboArguments.Count -eq 0) {
    throw 'Internal error: no Robocopy arguments were supplied.'
  }

  $line = (($RoboArguments | ForEach-Object { Quote-Arg $_ }) -join ' ')

  if ([string]::IsNullOrWhiteSpace($line)) {
    throw 'Internal error: Robocopy argument line is empty.'
  }

  return Start-Process -FilePath $exe -ArgumentList $line -NoNewWindow -PassThru
}

function Wait-Robo([System.Diagnostics.Process]$Process,[string]$Activity,[int]$DiskNumber,[string]$Serial) {
  $start = Get-Date
  $lastCheck = [datetime]::MinValue

  while (-not $Process.HasExited) {
    $now = Get-Date
    Write-Progress -Activity $Activity -Status ("Elapsed {0}" -f (Format-Duration ($now-$start))) -CurrentOperation $script:SourceRoot

    if (($now-$lastCheck).TotalSeconds -ge 30) {
      try { [void](Assert-ReadOnly $DiskNumber $Serial) }
      catch { try { $Process.Kill() } catch {}; throw }
      $lastCheck = $now
    }

    Start-Sleep -Milliseconds 750
    $Process.Refresh()
  }

  $Process.WaitForExit()
  Write-Progress -Activity $Activity -Completed
}

function Get-RoboSummary([string]$LogPath) {
  $tail = Get-Content -LiteralPath $LogPath -Encoding Unicode -Tail 160
  $rows = New-Object System.Collections.Generic.List[object]

  foreach ($line in $tail) {
    if ($line -match '^\s*[^:]+:\s+(?<Total>\d+)\s+(?<Copied>\d+)\s+(?<Skipped>\d+)\s+(?<Mismatch>\d+)\s+(?<Failed>\d+)\s+(?<Extras>\d+)\s*$') {
      $rows.Add([pscustomobject]@{
        Total=[UInt64]$matches.Total
        Skipped=[UInt64]$matches.Skipped
        Failed=[UInt64]$matches.Failed
      })
    }
  }

  if ($rows.Count -lt 3) { throw "Could not parse Robocopy summary from $LogPath" }

  return [pscustomobject]@{
    DirsTotal=$rows[0].Total
    DirsSkipped=$rows[0].Skipped
    DirsFailed=$rows[0].Failed
    FilesTotal=$rows[1].Total
    FilesSkipped=$rows[1].Skipped
    FilesFailed=$rows[1].Failed
    BytesTotal=$rows[2].Total
    BytesSkipped=$rows[2].Skipped
    BytesFailed=$rows[2].Failed
  }
}

function Watch-Raw(
  [System.Diagnostics.Process]$Process,
  [string]$LogPath,
  [UInt64]$ExpectedFiles,
  [UInt64]$ExpectedBytes,
  [int]$DiskNumber,
  [string]$Serial
) {
  while (-not (Test-Path -LiteralPath $LogPath)) {
    if ($Process.HasExited) { throw 'Robocopy exited before creating the raw log.' }
    Start-Sleep -Milliseconds 250
  }

  $stream = New-Object System.IO.FileStream($LogPath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
  $reader = New-Object System.IO.StreamReader($stream,[System.Text.Encoding]::Unicode,$true,65536)

  [UInt64]$seen=0
  [UInt64]$bytes=0
  [UInt64]$nextMark=100000
  $start=Get-Date
  $lastUi=[datetime]::MinValue
  $lastCheck=[datetime]::MinValue
  $current=''

  try {
    while ($true) {
      $line=$reader.ReadLine()

      if ($null -ne $line) {
        $m=$script:FileLineRegex.Match($line)

        if ($m.Success) {
          $seen++
          $bytes += [UInt64]$m.Groups['Size'].Value
          $current=$m.Groups['Path'].Value
          $now=Get-Date

          if (($now-$lastUi).TotalMilliseconds -ge 500) {
            $elapsed=$now-$start
            $pf=if($ExpectedFiles -gt 0){([double]$seen/[double]$ExpectedFiles)*100}else{0}
            $pb=if($ExpectedBytes -gt 0){([double]$bytes/[double]$ExpectedBytes)*100}else{0}
            $rate=if($elapsed.TotalSeconds -gt 0){[double]$seen/$elapsed.TotalSeconds}else{0}
            $eta=[TimeSpan]::Zero

            if($rate -gt 0 -and $seen -lt $ExpectedFiles){
              $eta=[TimeSpan]::FromSeconds(([double]($ExpectedFiles-$seen))/$rate)
            }

            $show=$current
            if($show.Length -gt 160){$show='...'+$show.Substring($show.Length-157)}

            $status='{0:N0}/{1:N0} files | files {2:N1}% | bytes {3:N1}% | {4:N1} files/s | ETA {5}' -f $seen,$ExpectedFiles,$pf,$pb,$rate,(Format-Duration $eta)
            Write-Progress -Activity 'Phase 2/3 - Raw listing' -Status $status -CurrentOperation $show -PercentComplete ([math]::Max(0,[math]::Min(100,$pf)))
            $lastUi=$now
          }

          if($seen -ge $nextMark){
            $msg='{0} | {1:N0}/{2:N0} files | {3:N1}% | {4:N2} GiB' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),$seen,$ExpectedFiles,(([double]$seen/[double]$ExpectedFiles)*100),([double]$bytes/1GB)
            Write-Host $msg
            Add-Content -LiteralPath $script:ProgressLog -Value $msg -Encoding UTF8
            while($nextMark -le $seen){$nextMark+=100000}
          }
        }
        continue
      }

      if($Process.HasExited){break}

      $now=Get-Date
      if(($now-$lastCheck).TotalSeconds -ge 30){
        try{[void](Assert-ReadOnly $DiskNumber $Serial)}
        catch{try{$Process.Kill()}catch{};throw}
        $lastCheck=$now
      }

      Start-Sleep -Milliseconds 200
      $Process.Refresh()
    }
  }
  finally {
    $reader.Dispose()
    $stream.Dispose()
    Write-Progress -Activity 'Phase 2/3 - Raw listing' -Completed
  }

  return [pscustomobject]@{FilesObserved=$seen;BytesObserved=$bytes}
}

function Parse-RawToCsv(
  [string]$RawLog,
  [string]$CsvPath,
  [string]$ParserErrors,
  [string]$Root,
  [string]$SourceId,
  [UInt64]$ExpectedFiles,
  [UInt64]$ExpectedBytes
) {
  $utf8bom=New-Object System.Text.UTF8Encoding($true)
  $reader=New-Object System.IO.StreamReader($RawLog,[System.Text.Encoding]::Unicode,$true,1048576)
  $writer=New-Object System.IO.StreamWriter($CsvPath,$false,$utf8bom,1048576)
  $errWriter=New-Object System.IO.StreamWriter($ParserErrors,$false,$utf8bom,65536)

  [UInt64]$n=0
  [UInt64]$sum=0
  [UInt64]$parseErr=0
  [UInt64]$nextMark=100000
  $start=Get-Date
  $lastUi=[datetime]::MinValue

  try {
    $writer.WriteLine('FileID;SourceID;FullPath;RelativePath;RelativeDirectory;FileName;BaseName;Extension;SizeBytes;LastWriteTime;LastWriteYear;TopLevelDirectory;PathLength;IsZeroByte')

    while(($line=$reader.ReadLine()) -ne $null){
      $m=$script:FileLineRegex.Match($line)

      if($m.Success){
        $n++
        $sizeText=$m.Groups['Size'].Value
        $size=[UInt64]$sizeText
        $sum+=$size

        $full=$m.Groups['Path'].Value
        $path=$full
        if($path.StartsWith('\\?\')){$path=$path.Substring(4)}

        if(-not $path.StartsWith($Root,[System.StringComparison]::OrdinalIgnoreCase)){
          $parseErr++
          $errWriter.WriteLine("SOURCE_MISMATCH|$line")
          continue
        }

        $rel=$path.Substring($Root.Length)
        $slash=$path.LastIndexOf('\')

        if($slash -ge 0){
          $dir=$path.Substring(0,$slash)
          $name=$path.Substring($slash+1)
        }else{
          $dir=''
          $name=$path
        }

        if($dir.Length -ge $Root.Length){$relDir=$dir.Substring($Root.Length).TrimStart('\')}else{$relDir=''}

        $first=$rel.IndexOf('\')
        if($first -gt 0){$top=$rel.Substring(0,$first)}else{$top='[ROOT]'}

        $dot=$name.LastIndexOf('.')
        if($dot -gt 0 -and $dot -lt ($name.Length-1)){
          $base=$name.Substring(0,$dot)
          $ext=$name.Substring($dot).ToLowerInvariant()
        }else{
          $base=$name
          $ext=''
        }

        $date=$m.Groups['Date'].Value
        $time=$m.Groups['Time'].Value
        $stamp=$date.Replace('/','-')+' '+$time
        $year=$date.Substring(0,4)
        $id='F'+$n.ToString('D9')
        $zero=if($size -eq 0){'1'}else{'0'}

        $fields=@(
          (Csv $id),(Csv $SourceId),(Csv $path),(Csv $rel),(Csv $relDir),
          (Csv $name),(Csv $base),(Csv $ext),$sizeText,(Csv $stamp),$year,
          (Csv $top),$path.Length,$zero
        )

        $writer.WriteLine(($fields -join ';'))

        $now=Get-Date
        if(($now-$lastUi).TotalMilliseconds -ge 500){
          $elapsed=$now-$start
          $pf=if($ExpectedFiles -gt 0){([double]$n/[double]$ExpectedFiles)*100}else{0}
          $pb=if($ExpectedBytes -gt 0){([double]$sum/[double]$ExpectedBytes)*100}else{0}
          $rate=if($elapsed.TotalSeconds -gt 0){[double]$n/$elapsed.TotalSeconds}else{0}
          $eta=[TimeSpan]::Zero
          if($rate -gt 0 -and $n -lt $ExpectedFiles){$eta=[TimeSpan]::FromSeconds(([double]($ExpectedFiles-$n))/$rate)}
          $status='{0:N0}/{1:N0} files | {2:N1}% | bytes {3:N1}% | ETA {4}' -f $n,$ExpectedFiles,$pf,$pb,(Format-Duration $eta)
          Write-Progress -Activity 'Phase 3/3 - Master CSV' -Status $status -CurrentOperation $path -PercentComplete ([math]::Max(0,[math]::Min(100,$pf)))
          $lastUi=$now
        }

        if($n -ge $nextMark){
          $msg='{0} | CSV {1:N0}/{2:N0} files | {3:N1}% | {4:N2} GiB' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),$n,$ExpectedFiles,(([double]$n/[double]$ExpectedFiles)*100),([double]$sum/1GB)
          Write-Host $msg
          Add-Content -LiteralPath $script:ProgressLog -Value $msg -Encoding UTF8
          while($nextMark -le $n){$nextMark+=100000}
        }
      }
      elseif($line -match [regex]::Escape($Root) -and $line -match '\d{4}/\d{2}/\d{2}' -and $line -match '\d{2}:\d{2}:\d{2}'){
        $parseErr++
        $errWriter.WriteLine("UNPARSED|$line")
      }
    }
  }
  finally {
    $writer.Flush();$writer.Dispose()
    $errWriter.Flush();$errWriter.Dispose()
    $reader.Dispose()
    Write-Progress -Activity 'Phase 3/3 - Master CSV' -Completed
  }

  return [pscustomobject]@{Records=$n;Bytes=$sum;ParserErrors=$parseErr}
}

function Export-RoboErrors([string[]]$Logs,[string]$Output) {
  $utf8bom=New-Object System.Text.UTF8Encoding($true)
  $writer=New-Object System.IO.StreamWriter($Output,$false,$utf8bom,65536)
  [UInt64]$count=0

  try{
    foreach($log in $Logs){
      if(-not(Test-Path -LiteralPath $log)){continue}
      $reader=New-Object System.IO.StreamReader($log,[System.Text.Encoding]::Unicode,$true,65536)
      try{
        while(($line=$reader.ReadLine()) -ne $null){
          if($line -match '(?i)\b(?:FEHLER|ERROR)\s+\d+\b' -or $line -match '(?i)Zugriff.+verweigert' -or $line -match '(?i)Access.+denied'){
            $count++
            $writer.WriteLine("[$([System.IO.Path]::GetFileName($log))] $line")
          }
        }
      }finally{$reader.Dispose()}
    }
  }finally{$writer.Flush();$writer.Dispose()}

  return $count
}

function Write-Manifest([string]$Status,[string]$Validation,[hashtable]$Extra=@{}) {
  $lines=New-Object System.Collections.Generic.List[string]
  $lines.Add('FILE INVENTATOR V1 - MANIFEST')
  $lines.Add('=============================')
  $lines.Add('')
  $lines.Add("ScriptVersion=$script:ScriptVersion")
  $lines.Add("RunID=$script:RunID")
  $lines.Add("Status=$Status")
  $lines.Add("Validation=$Validation")
  $lines.Add("Started=$($script:RunStarted.ToString('yyyy-MM-dd HH:mm:ss'))")
  $lines.Add("Updated=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
  $lines.Add('')
  $lines.Add("SourceRoot=$script:SourceRoot")
  $lines.Add("SourceID=$script:SourceID")
  $lines.Add("DiskNumber=$script:DiskNumber")
  $lines.Add("DiskFriendlyName=$script:DiskFriendlyName")
  $lines.Add("DiskSerialNumber=$script:DiskSerialNumber")
  $lines.Add("DiskUniqueId=$script:DiskUniqueId")
  $lines.Add("DiskPartitionStyle=$script:DiskPartitionStyle")
  $lines.Add("DiskSizeBytes=$script:DiskSizeBytes")
  $lines.Add("DiskReadOnlyAtStart=$script:DiskReadOnlyAtStart")
  $lines.Add("VolumeLabel=$script:VolumeLabel")
  $lines.Add("VolumeFileSystem=$script:VolumeFileSystem")
  $lines.Add("VolumeSizeBytes=$script:VolumeSizeBytes")
  $lines.Add('')
  $lines.Add('Enumeration=ROBOCOPY /L /B /E /XJ /SL')
  $lines.Add('ContentHashesCalculated=False')
  $lines.Add('JunctionTargetsFollowed=False')
  $lines.Add('SymbolicLinkTargetsFollowed=False')
  $lines.Add('DeletedFilesRecovered=False')
  $lines.Add('UnallocatedSpaceScanned=False')
  $lines.Add('')
  $lines.Add("OutputDirectory=$script:OutputDir")
  $lines.Add("MasterCsv=$script:MasterCsv")
  $lines.Add("RawListing=$script:RawLog")
  $lines.Add("PrecountLog=$script:PrecountLog")
  $lines.Add("ErrorsLog=$script:ErrorsLog")
  $lines.Add("ParserErrorsLog=$script:ParserErrorsLog")

  foreach($key in ($Extra.Keys|Sort-Object)){$lines.Add("$key=$($Extra[$key])")}
  $lines|Set-Content -LiteralPath $script:ManifestPath -Encoding UTF8
}

# Initial safety checks
$SourceRoot=Normalize-Root $SourceRoot
$ExpectedSerial=$ExpectedSerial.Trim()
$disk=Get-CheckedDisk $SourceRoot $ExpectedSerial
$letter=$SourceRoot.Substring(0,1)
$volume=Get-Volume -DriveLetter $letter -ErrorAction Stop

$DiskNumber=[int]$disk.Number
$DiskFriendlyName=[string]$disk.FriendlyName
$DiskSerialNumber=([string]$disk.SerialNumber).Trim()
$DiskUniqueId=[string]$disk.UniqueId
$DiskPartitionStyle=[string]$disk.PartitionStyle
$DiskSizeBytes=[UInt64]$disk.Size
$DiskReadOnlyAtStart=[bool]$disk.IsReadOnly
$VolumeLabel=[string]$volume.FileSystemLabel
$VolumeFileSystem=[string]$volume.FileSystem
$VolumeSizeBytes=[UInt64]$volume.Size
$SourceID='DISK-'+$DiskSerialNumber

$Desktop=[Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
if([string]::IsNullOrWhiteSpace($Desktop)){throw 'Could not resolve Desktop path.'}

$desktopRoot=[System.IO.Path]::GetPathRoot($Desktop)
if($desktopRoot.TrimEnd('\').Equals($SourceRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: Desktop output and source volume are on the same drive.'
}

$OutputDir=Join-Path $Desktop ($ScriptName+'_'+$RunID)
$i=2
while(Test-Path -LiteralPath $OutputDir){
  $OutputDir=Join-Path $Desktop (($ScriptName+'_'+$RunID)+'_{0:D2}' -f $i)
  $i++
}
New-Item -ItemType Directory -Path $OutputDir|Out-Null

$DummyTarget=Join-Path $OutputDir '_robocopy_dummy_target'
$PrecountLog=Join-Path $OutputDir ($ScriptName+'_precount.txt')
$RawLog=Join-Path $OutputDir ($ScriptName+'_raw.txt')
$MasterCsv=Join-Path $OutputDir ($ScriptName+'_master.csv')
$ErrorsLog=Join-Path $OutputDir ($ScriptName+'_errors.txt')
$ParserErrorsLog=Join-Path $OutputDir ($ScriptName+'_parser-errors.txt')
$ProgressLog=Join-Path $OutputDir ($ScriptName+'_progress.txt')
$ManifestPath=Join-Path $OutputDir ($ScriptName+'_manifest.txt')

New-Item -ItemType Directory -Path $DummyTarget|Out-Null
New-Item -ItemType File -Path $ProgressLog|Out-Null
Write-Manifest 'IN_PROGRESS' 'PENDING'

Write-Host
Write-Host 'FILE INVENTATOR V1'
Write-Host '=================='
Write-Host "Source       : $SourceRoot"
Write-Host "Disk         : $DiskFriendlyName"
Write-Host "Serial       : $DiskSerialNumber"
Write-Host "Read-only    : $DiskReadOnlyAtStart"
Write-Host "Output       : $OutputDir"
Write-Host "Content hash : NO"
Write-Host

try{
  # Phase 1
  Write-Host 'Phase 1/3: Pre-count with backup rights ...'

  $preArgs=@($SourceRoot,$DummyTarget,'/L','/B','/E','/XJ','/SL','/R:0','/W:0','/BYTES','/NFL','/NDL','/NP',"/UNILOG:$PrecountLog")
  $pre=Start-Robo $preArgs
  Wait-Robo $pre 'Phase 1/3 - Pre-count' $DiskNumber $ExpectedSerial

  $preExit=[int]$pre.ExitCode
  $preSum=Get-RoboSummary $PrecountLog

  if($preExit -ge 8){throw "Robocopy precount exit code $preExit"}
  if($preSum.FilesFailed -gt 0 -or $preSum.DirsFailed -gt 0){throw 'Robocopy precount reported failures.'}

  [UInt64]$ExpectedFiles=$preSum.FilesTotal
  [UInt64]$ExpectedBytes=$preSum.BytesTotal
  if($ExpectedFiles -eq 0){throw 'No files found.'}

  Write-Host ("Expected files : {0:N0}" -f $ExpectedFiles)
  Write-Host ("Expected bytes : {0:N0}" -f $ExpectedBytes)
  Write-Host ("Expected GiB   : {0:N3}" -f ([double]$ExpectedBytes/1GB))
  Write-Host

  $estimate=([double]$ExpectedFiles*1500.0)+200MB
  $outDrive=New-Object System.IO.DriveInfo($desktopRoot)
  if([double]$outDrive.AvailableFreeSpace -lt $estimate){throw 'Insufficient free space at output location.'}

  [void](Assert-ReadOnly $DiskNumber $ExpectedSerial)

  # Phase 2
  Write-Host 'Phase 2/3: Creating Unicode raw listing ...'

  $fullArgs=@($SourceRoot,$DummyTarget,'/L','/B','/E','/XJ','/SL','/R:0','/W:0','/BYTES','/FP','/TS','/NP','/NDL',"/UNILOG:$RawLog")
  $full=Start-Robo $fullArgs
  $live=Watch-Raw $full $RawLog $ExpectedFiles $ExpectedBytes $DiskNumber $ExpectedSerial

  $full.WaitForExit()
  $fullExit=[int]$full.ExitCode
  $fullSum=Get-RoboSummary $RawLog

  if($fullExit -ge 8){throw "Robocopy full-run exit code $fullExit"}
  if($fullSum.FilesFailed -gt 0 -or $fullSum.DirsFailed -gt 0){throw 'Robocopy full run reported failures.'}

  Write-Host ("Raw listing completed: {0:N0} files." -f $fullSum.FilesTotal)

  # Phase 3
  Write-Host 'Phase 3/3: Building master CSV and validating ...'

  $parsed=Parse-RawToCsv $RawLog $MasterCsv $ParserErrorsLog $SourceRoot $SourceID $ExpectedFiles $ExpectedBytes
  $roboErrors=Export-RoboErrors @($PrecountLog,$RawLog) $ErrorsLog
  $diskEnd=Assert-ReadOnly $DiskNumber $ExpectedSerial

  $checks=[ordered]@{
    PrecountVsFullFiles=($preSum.FilesTotal -eq $fullSum.FilesTotal)
    PrecountVsFullBytes=($preSum.BytesTotal -eq $fullSum.BytesTotal)
    FullVsCsvFiles=($fullSum.FilesTotal -eq $parsed.Records)
    FullVsCsvBytes=($fullSum.BytesTotal -eq $parsed.Bytes)
    PrecountFailures0=(($preSum.FilesFailed+$preSum.DirsFailed) -eq 0)
    FullFailures0=(($fullSum.FilesFailed+$fullSum.DirsFailed) -eq 0)
    ParserErrors0=($parsed.ParserErrors -eq 0)
    PrecountExitOK=($preExit -lt 8)
    FullExitOK=($fullExit -lt 8)
    SourceStillReadOnly=([bool]$diskEnd.IsReadOnly)
  }

  $pass=$true
  foreach($v in $checks.Values){if(-not $v){$pass=$false;break}}

  if($pass){$status='COMPLETE';$validation='PASS'}
  else{$status='COMPLETE_WITH_VALIDATION_FAILURE';$validation='FAIL'}

  Write-Manifest $status $validation @{
    PrecountExitCode=$preExit
    FullRunExitCode=$fullExit
    PrecountDirs=$preSum.DirsTotal
    PrecountFiles=$preSum.FilesTotal
    PrecountBytes=$preSum.BytesTotal
    FullDirs=$fullSum.DirsTotal
    FullFiles=$fullSum.FilesTotal
    FullBytes=$fullSum.BytesTotal
    LiveFilesObserved=$live.FilesObserved
    LiveBytesObserved=$live.BytesObserved
    CsvRecords=$parsed.Records
    CsvBytes=$parsed.Bytes
    ParserErrors=$parsed.ParserErrors
    RobocopyErrorLines=$roboErrors
    Check_PrecountVsFullFiles=$checks.PrecountVsFullFiles
    Check_PrecountVsFullBytes=$checks.PrecountVsFullBytes
    Check_FullVsCsvFiles=$checks.FullVsCsvFiles
    Check_FullVsCsvBytes=$checks.FullVsCsvBytes
    Check_SourceStillReadOnly=$checks.SourceStillReadOnly
    ExcelRowLimitExceeded=($parsed.Records -gt 1048575)
    Finished=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  }

  try{Remove-Item -LiteralPath $DummyTarget -Force -Recurse -ErrorAction SilentlyContinue}catch{}

  Write-Host
  Write-Host 'DONE'
  Write-Host "Status      : $status"
  Write-Host "Validation  : $validation"
  Write-Host ("Files       : {0:N0}" -f $parsed.Records)
  Write-Host ("GiB         : {0:N3}" -f ([double]$parsed.Bytes/1GB))
  Write-Host ("Parser errs : {0:N0}" -f $parsed.ParserErrors)
  Write-Host ("Robo errors : {0:N0}" -f $roboErrors)
  Write-Host "Results     : $OutputDir"

  if($parsed.Records -gt 1048575){Write-Warning 'CSV exceeds one Excel worksheet row limit; the CSV itself remains complete.'}
  if($validation -ne 'PASS'){Write-Warning 'Completeness validation did not fully pass. Review manifest and error logs.'}
}
catch{
  $msg=$_.Exception.Message

  try{
    Write-Manifest 'FAILED' 'FAIL' @{
      FailureMessage=$msg
      FailedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
  }catch{}

  Write-Error "file-inventator-v1 aborted: $msg"
  Write-Host "Partial output remains at: $OutputDir"
  throw
}
