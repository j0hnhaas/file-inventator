#requires -Version 5.1

<#
.SYNOPSIS
  Optional utility: create a UDF ISO image from a completed recovery-audit
  sample folder.

.DESCRIPTION
  recovery-audit-iso.ps1 packages a local sample folder into an ISO file.
  It is intentionally separate from recovery-audit-copy.ps1: creating an ISO
  is optional and requires its own explicit command.

  The script uses Windows Image Mastering API v2 (IMAPI2FS) and creates a
  single-session UDF image suitable for DVD-sized data sets. It does not alter
  the source folder.

  If copy-summary.txt is present in SourceFolder, the script requires
  Status=COMPLETE and FailedFiles=0 unless -AllowIncompleteCopy is supplied.

  The resulting image size is checked against MaxIsoBytes before it is written.
  Default MaxIsoBytes is 4,700,000,000 bytes.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-iso.ps1" -SourceFolder "C:\RecoveryAuditSample" -IsoPath "C:\RecoveryAuditSample.iso" -VolumeLabel "RECOVERY_AUDIT"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$SourceFolder,

  [Parameter(Mandatory=$true)]
  [string]$IsoPath,

  [Parameter()]
  [ValidatePattern('^[A-Za-z0-9_\-]{1,32}$')]
  [string]$VolumeLabel='RECOVERY_AUDIT',

  [Parameter()]
  [UInt64]$MaxIsoBytes=4700000000,

  [Parameter()]
  [switch]$AllowIncompleteCopy
)

$ErrorActionPreference='Stop'
$ScriptVersion='1.0.4'
$Started=Get-Date

function Format-Duration([TimeSpan]$Span) {
  $h=[int][math]::Floor($Span.TotalHours)
  return ('{0:00}:{1:00}:{2:00}' -f $h,$Span.Minutes,$Span.Seconds)
}

function Convert-ToExtendedPath([string]$Path) {
  $full=[System.IO.Path]::GetFullPath($Path)
  if($full.StartsWith('\\?\')){return $full}
  if($full.StartsWith('\\')){
    return '\\?\UNC\'+$full.Substring(2)
  }
  return '\\?\'+$full
}

function Get-FileLengthExtended([string]$Path) {
  $stream=New-Object System.IO.FileStream((Convert-ToExtendedPath $Path),[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
  try{return [UInt64]$stream.Length}
  finally{$stream.Dispose()}
}

function Get-KeyValueFile([string]$Path) {
  $map=@{}
  foreach($line in Get-Content -LiteralPath $Path -Encoding UTF8){
    if($line -match '^(?<Key>[^=]+)=(?<Value>.*)$'){
      $map[$matches.Key]=$matches.Value
    }
  }
  return $map
}

$SourceFolder=[System.IO.Path]::GetFullPath($SourceFolder)
$IsoPath=[System.IO.Path]::GetFullPath($IsoPath)

if(-not(Test-Path -LiteralPath $SourceFolder -PathType Container)){
  throw "Source folder not found: $SourceFolder"
}

if([System.IO.Path]::GetExtension($IsoPath) -ne '.iso'){
  throw "IsoPath must end in .iso: $IsoPath"
}

if(Test-Path -LiteralPath $IsoPath){
  throw "ISO already exists. Refusing to overwrite: $IsoPath"
}

$isoParent=[System.IO.Path]::GetDirectoryName($IsoPath)
if([string]::IsNullOrWhiteSpace($isoParent)){
  throw "Could not resolve ISO parent directory: $IsoPath"
}

if(-not(Test-Path -LiteralPath $isoParent -PathType Container)){
  New-Item -ItemType Directory -Path $isoParent -Force | Out-Null
}

$sourceRoot=([System.IO.Path]::GetPathRoot($SourceFolder)).TrimEnd('\')
$isoRoot=([System.IO.Path]::GetPathRoot($IsoPath)).TrimEnd('\')

$sourcePrefix=$SourceFolder.TrimEnd('\')+'\'
if($IsoPath.StartsWith($sourcePrefix,[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: ISO output must not be created inside SourceFolder.'
}

$reconciledSummaryPath=Join-Path $SourceFolder 'copy-summary-reconciled.txt'
$copySummaryPath=if(Test-Path -LiteralPath $reconciledSummaryPath -PathType Leaf){
  $reconciledSummaryPath
}else{
  Join-Path $SourceFolder 'copy-summary.txt'
}

if(Test-Path -LiteralPath $copySummaryPath -PathType Leaf){
  $copySummary=Get-KeyValueFile $copySummaryPath

  if(-not $AllowIncompleteCopy){
    if(-not $copySummary.ContainsKey('Status') -or $copySummary['Status'] -ne 'COMPLETE'){
      throw "PRECHECK ABORT: $([System.IO.Path]::GetFileName($copySummaryPath)) does not report Status=COMPLETE. Use -AllowIncompleteCopy only for deliberate exception handling."
    }

    if(-not $copySummary.ContainsKey('FailedFiles') -or [UInt64]$copySummary['FailedFiles'] -ne 0){
      throw "PRECHECK ABORT: $([System.IO.Path]::GetFileName($copySummaryPath)) reports failed or unknown copy count. Use -AllowIncompleteCopy only for deliberate exception handling."
    }
  }
}

Write-Host
Write-Host 'RECOVERY AUDIT ISO - PREFLIGHT'
Write-Host '=============================='
Write-Host "Source folder : $SourceFolder"
Write-Host "ISO path      : $IsoPath"
Write-Host "Volume label  : $VolumeLabel"
Write-Host ("Maximum bytes : {0:N0}" -f $MaxIsoBytes)
Write-Host

[UInt64]$SourceFiles=0
[UInt64]$SourceBytes=0
$scanStarted=Get-Date
$lastUi=[datetime]::MinValue

$extendedSource=Convert-ToExtendedPath $SourceFolder
$files=[System.IO.Directory]::EnumerateFiles($extendedSource,'*',[System.IO.SearchOption]::AllDirectories)

foreach($file in $files){
  $SourceFiles++
  $SourceBytes+=Get-FileLengthExtended $file

  $displayPath=$file
  if($displayPath.StartsWith('\\?\')){$displayPath=$displayPath.Substring(4)}

  $now=Get-Date
  if(($now-$lastUi).TotalMilliseconds -ge 250){
    $elapsed=$now-$scanStarted
    $rate=if($elapsed.TotalSeconds -gt 0){[double]$SourceFiles/$elapsed.TotalSeconds}else{0}
    $status=('{0:N0} files | {1:N3} GB | {2:N1} files/s' -f $SourceFiles,([double]$SourceBytes/1000000000),$rate)
    Write-Progress -Activity 'Recovery Audit ISO - scanning source folder' -Status $status -CurrentOperation $displayPath
    $lastUi=$now
  }
}

Write-Progress -Activity 'Recovery Audit ISO - scanning source folder' -Completed

if($SourceFiles -eq 0){
  throw 'Source folder contains no files.'
}

Write-Host ("Source files  : {0:N0}" -f $SourceFiles)
Write-Host ("Source bytes  : {0:N0} ({1:N3} decimal GB)" -f $SourceBytes,([double]$SourceBytes/1000000000))
Write-Host
Write-Host 'Building UDF file-system image ...'

Write-Progress -Activity 'Recovery Audit ISO - building file-system image' -Status 'Adding directory tree to UDF image'

$fsi=New-Object -ComObject IMAPI2FS.MsftFileSystemImage
try{
  # IMAPI_MEDIA_TYPE_DVDPLUSR = 6. Use a writable single-layer DVD profile;
  # DVDROM (4) is read-only media and can be rejected by IMAPI2FS when used
  # as the target profile for a burn image.
  $fsi.ChooseImageDefaultsForMediaType([int]6)

  # FsiFileSystemUDF = 4. UDF is used to preserve long Windows file names
  # more reliably than ISO9660/Joliet-only layouts.
  $fsi.FileSystemsToCreate=4
  $fsi.VolumeName=$VolumeLabel
  $fsi.StageFiles=$false

  # Add the contents of SourceFolder at the ISO root, not the base folder itself.
  $fsi.Root.AddTree((Convert-ToExtendedPath $SourceFolder),$false)

  Write-Progress -Activity 'Recovery Audit ISO - building file-system image' -Status 'Finalising image layout'

  $result=$fsi.CreateResultImage()
  [UInt64]$TotalBlocks=[UInt64]$result.TotalBlocks
  [UInt64]$BlockSize=[UInt64]$result.BlockSize
  [UInt64]$ImageBytes=$TotalBlocks*$BlockSize

  Write-Progress -Activity 'Recovery Audit ISO - building file-system image' -Completed

  if($ImageBytes -gt $MaxIsoBytes){
    throw ("ISO image would be {0:N0} bytes, exceeding MaxIsoBytes={1:N0}." -f $ImageBytes,$MaxIsoBytes)
  }

  $isoDriveName=([System.IO.Path]::GetPathRoot($IsoPath)).Substring(0,1)
  $drive=Get-PSDrive -Name $isoDriveName -ErrorAction Stop
  if([UInt64]$drive.Free -lt ($ImageBytes+100000000)){
    throw ("Destination has {0:N0} free bytes; at least {1:N0} are required." -f [UInt64]$drive.Free,($ImageBytes+100000000))
  }

  $imageStream=$result.ImageStream
  try{
    $comStream=[System.Runtime.InteropServices.ComTypes.IStream]$imageStream
  }
  catch{
    throw "Could not expose IMAPI ImageStream as System.Runtime.InteropServices.ComTypes.IStream: $($_.Exception.Message)"
  }

  if($null -eq $comStream){
    throw 'Could not expose IMAPI ImageStream as System.Runtime.InteropServices.ComTypes.IStream.'
  }

  $outStream=New-Object System.IO.FileStream($IsoPath,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None,1048576)
  $buffer=New-Object byte[] 1048576
  $readPtr=[System.Runtime.InteropServices.Marshal]::AllocCoTaskMem(4)
  [UInt64]$Written=0
  $writeStarted=Get-Date
  $lastUi=[datetime]::MinValue

  try{
    while($true){
      [System.Runtime.InteropServices.Marshal]::WriteInt32($readPtr,0)
      $comStream.Read($buffer,$buffer.Length,$readPtr)
      $read=[System.Runtime.InteropServices.Marshal]::ReadInt32($readPtr)
      if($read -le 0){break}

      $outStream.Write($buffer,0,$read)
      $Written+=[UInt64]$read

      $now=Get-Date
      if(($now-$lastUi).TotalMilliseconds -ge 250 -or $Written -ge $ImageBytes){
        $percent=if($ImageBytes -gt 0){([double]$Written/[double]$ImageBytes)*100}else{100}
        $elapsed=$now-$writeStarted
        $rate=if($elapsed.TotalSeconds -gt 0){[double]$Written/$elapsed.TotalSeconds}else{0}
        $eta=[TimeSpan]::Zero

        if($rate -gt 0 -and $Written -lt $ImageBytes){
          $eta=[TimeSpan]::FromSeconds(([double]($ImageBytes-$Written))/$rate)
        }

        $status=('{0:N3}/{1:N3} GB | {2:N1}% | {3:N1} MB/s | ETA {4}' -f ([double]$Written/1000000000),([double]$ImageBytes/1000000000),$percent,($rate/1000000),(Format-Duration $eta))
        $progressPercent=[math]::Max(0,[math]::Min(100,$percent))
        Write-Progress -Activity 'Recovery Audit ISO - writing image' -Status $status -CurrentOperation $IsoPath -PercentComplete $progressPercent
        $lastUi=$now
      }
    }

    $outStream.Flush()
  }
  finally{
    $outStream.Dispose()
    if($readPtr -ne [IntPtr]::Zero){
      [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($readPtr)
    }
    Write-Progress -Activity 'Recovery Audit ISO - writing image' -Completed
  }

  if($Written -ne $ImageBytes){
    if(Test-Path -LiteralPath $IsoPath){Remove-Item -LiteralPath $IsoPath -Force -ErrorAction SilentlyContinue}
    throw "ISO write length mismatch. Expected $ImageBytes bytes, wrote $Written bytes."
  }

  $isoInfo=Get-Item -LiteralPath $IsoPath -Force -ErrorAction Stop
  if([UInt64]$isoInfo.Length -ne $ImageBytes){
    if(Test-Path -LiteralPath $IsoPath){Remove-Item -LiteralPath $IsoPath -Force -ErrorAction SilentlyContinue}
    throw "ISO file size mismatch after close. Expected $ImageBytes bytes, found $($isoInfo.Length) bytes."
  }

  $Finished=Get-Date

  Write-Host
  Write-Host 'RECOVERY AUDIT ISO - DONE'
  Write-Host '========================='
  Write-Host ("Source files   : {0:N0}" -f $SourceFiles)
  Write-Host ("Source GB      : {0:N3}" -f ([double]$SourceBytes/1000000000))
  Write-Host ("ISO bytes      : {0:N0}" -f $ImageBytes)
  Write-Host ("ISO GB         : {0:N3}" -f ([double]$ImageBytes/1000000000))
  Write-Host ("Duration       : {0}" -f (Format-Duration ($Finished-$Started)))
  Write-Host "File system    : UDF"
  Write-Host "ISO            : $IsoPath"
}
catch{
  if(Test-Path -LiteralPath $IsoPath -PathType Leaf){
    try{Remove-Item -LiteralPath $IsoPath -Force -ErrorAction Stop}catch{}
  }
  throw
}
finally{
  if($null -ne $result){
    try{[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($result)}catch{}
  }
  if($null -ne $fsi){
    try{[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fsi)}catch{}
  }
}
