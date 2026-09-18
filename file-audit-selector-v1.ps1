#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
  Builds a size-limited audit sample plan from a file-inventator CSV.

.DESCRIPTION
  Planning mode only: no source files are copied or opened.
  The inventory CSV is read as a stream. The mounted source disk must match
  ExpectedSerial and must report IsReadOnly=True throughout the run.

  Priority:
    1) Desktop and subfolders near the documented failure event
    2) Desktop in the preceding four months
    3) Other personal files in the preceding four months
    4) Historical Desktop files
    5) Other historical files
    6) Legacy controls

  Oversize, zero-byte, intervention-window, and post-intervention files are
  reported separately and are not automatically selected.

  Default sample limit: 3,900,000,000 bytes.
  Default oversize threshold: 500,000,000 bytes.

.EXAMPLE
  .\file-audit-selector-v1.ps1 -InventoryCsv ".\file-inventator-v1_master.csv" -SourceRoot "X:\" -ExpectedSerial "SERIAL_NUMBER" -UserRoot "X:\Users\PROFILE\" -FirstDocumentedFailureTime "2026-07-26 20:32:24" -InterventionEndTime "2026-07-27 04:00:00"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$InventoryCsv,
  [Parameter(Mandatory=$true)][ValidatePattern('^[A-Za-z]:\\?$')][string]$SourceRoot,
  [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ExpectedSerial,
  [Parameter(Mandatory=$true)][string]$UserRoot,
  [Parameter()][string[]]$AdditionalRoots=@(),
  [Parameter(Mandatory=$true)][datetime]$FirstDocumentedFailureTime,
  [Parameter(Mandatory=$true)][datetime]$InterventionEndTime,
  [Parameter()][UInt64]$MaxSampleBytes=3900000000,
  [Parameter()][UInt64]$OversizeThresholdBytes=500000000,
  [Parameter()][int]$HistoricalStartYear=2020,
  [Parameter()][string]$PlannedSampleFolderName='!AnalyseSampleRecovery',
  [Parameter()][switch]$IncludeNoisePaths
)

$ErrorActionPreference='Stop'
$ScriptVersion='1.0'
$RunID=(Get-Date).ToString('yyyyMMdd_HHmmss')
Add-Type -AssemblyName Microsoft.VisualBasic

function Normalize-Root([string]$Path) {
  $p=[System.IO.Path]::GetFullPath($Path)
  if(-not $p.EndsWith('\')){$p+='\'}
  return $p
}

function Csv([AllowNull()][string]$Value) {
  if($null -eq $Value){return '""'}
  return '"' + $Value.Replace('"','""') + '"'
}

function Get-CheckedDisk([string]$Root,[string]$Serial) {
  $letter=$Root.Substring(0,1)
  $partition=Get-Partition -DriveLetter $letter -ErrorAction Stop
  $disk=$partition | Get-Disk -ErrorAction Stop
  $actual=([string]$disk.SerialNumber).Trim()
  if($actual -ne $Serial.Trim()){throw "SAFETY ABORT: source serial '$actual' does not match expected serial '$Serial'."}
  if(-not $disk.IsReadOnly){throw "SAFETY ABORT: source disk $($disk.Number) is NOT read-only."}
  return $disk
}

function Assert-ReadOnly([int]$DiskNumber,[string]$Serial) {
  $disk=Get-Disk -Number $DiskNumber -ErrorAction Stop
  if((([string]$disk.SerialNumber).Trim()) -ne $Serial.Trim()){throw 'SAFETY ABORT: source disk identity changed.'}
  if(-not $disk.IsReadOnly){throw 'SAFETY ABORT: source disk is no longer read-only.'}
  return $disk
}

function Get-Group([string]$Ext) {
  switch($Ext.ToLowerInvariant()){
    {$_ -in @('.docx','.xlsx','.xls','.doc','.pptx','.rtf','.txt','.md','.pdf')}{return 'OfficeText'}
    {$_ -in @('.psd','.ai','.indd')}{return 'Design'}
    {$_ -in @('.jpg','.jpeg','.png','.tif','.tiff')}{return 'Image'}
    {$_ -in @('.wav','.mp3','.m4a','.aac','.aup3')}{return 'Audio'}
    {$_ -in @('.csv','.json','.py')}{return 'DataCode'}
    {$_ -in @('.zip','.rar')}{return 'Archive'}
    default{return 'Other'}
  }
}

function Get-TypeRank([string]$Ext) {
  switch($Ext.ToLowerInvariant()){
    {$_ -in @('.docx','.xlsx','.xls','.doc','.pptx','.rtf','.txt','.md','.psd','.ai','.indd','.aup3','.csv','.py')}{return 1}
    {$_ -in @('.pdf','.json')}{return 2}
    {$_ -in @('.jpg','.jpeg','.png','.tif','.tiff','.wav','.mp3','.m4a','.aac')}{return 3}
    {$_ -in @('.zip','.rar')}{return 4}
    default{return 9}
  }
}

function Get-TimeBand([datetime]$Stamp) {
  $a=$script:FirstDocumentedFailureTime
  if($Stamp -ge $a -and $Stamp -le $script:InterventionEndTime){return 'INTERVENTION_WINDOW'}
  if($Stamp -gt $script:InterventionEndTime){return 'AFTER_INTERVENTION'}
  if($Stamp -ge $a.AddHours(-1)){return 'LAST_1_HOUR'}
  if($Stamp -ge $a.AddHours(-6)){return 'LAST_6_HOURS'}
  if($Stamp -ge $a.Date){return 'EVENT_DAY_EARLIER'}
  if($Stamp -ge $a.Date.AddDays(-1)){return 'DAY_BEFORE'}
  if($Stamp -ge $a.AddDays(-7)){return 'LAST_7_DAYS'}
  if($Stamp -ge $a.AddDays(-30)){return 'LAST_30_DAYS'}
  if($Stamp -ge $a.AddMonths(-4)){return 'LAST_4_MONTHS'}
  if($Stamp.Year -ge $script:HistoricalStartYear){return 'HISTORICAL'}
  return 'LEGACY'
}

function Get-TimeRank([string]$Band) {
  switch($Band){
    'LAST_1_HOUR'{return 1}
    'LAST_6_HOURS'{return 2}
    'EVENT_DAY_EARLIER'{return 3}
    'DAY_BEFORE'{return 4}
    'LAST_7_DAYS'{return 5}
    'LAST_30_DAYS'{return 6}
    'LAST_4_MONTHS'{return 7}
    'HISTORICAL'{return 8}
    default{return 9}
  }
}

function Get-Bucket([bool]$Desktop,[string]$Band) {
  if($Desktop -and $Band -in @('LAST_1_HOUR','LAST_6_HOURS','EVENT_DAY_EARLIER','DAY_BEFORE','LAST_7_DAYS')){return 'A_CRITICAL_DESKTOP'}
  if($Desktop -and $Band -in @('LAST_30_DAYS','LAST_4_MONTHS')){return 'B_DESKTOP_4MONTHS'}
  if((-not $Desktop) -and $Band -in @('LAST_1_HOUR','LAST_6_HOURS','EVENT_DAY_EARLIER','DAY_BEFORE','LAST_7_DAYS','LAST_30_DAYS','LAST_4_MONTHS')){return 'C_OTHER_4MONTHS'}
  if($Desktop -and $Band -eq 'HISTORICAL'){return 'D_DESKTOP_HISTORICAL'}
  if((-not $Desktop) -and $Band -eq 'HISTORICAL'){return 'E_OTHER_HISTORICAL'}
  return 'F_LEGACY_CONTROL'
}

function Is-Noise([string]$Path) {
  if($script:IncludeNoisePaths){return $false}
  foreach($p in @('\AppData\','\node_modules\','\site-packages\','\__pycache__\','\.git\','\.cache\','\venv\','\.venv\','\OneDriveTemp\')){
    if($Path.IndexOf($p,[System.StringComparison]::OrdinalIgnoreCase) -ge 0){return $true}
  }
  return $false
}

$InventoryCsv=[System.IO.Path]::GetFullPath($InventoryCsv)
if(-not(Test-Path -LiteralPath $InventoryCsv)){throw "Inventory CSV not found: $InventoryCsv"}
$SourceRoot=Normalize-Root $SourceRoot
$UserRoot=Normalize-Root $UserRoot
$ExpectedSerial=$ExpectedSerial.Trim()
if(-not $UserRoot.StartsWith($SourceRoot,[System.StringComparison]::OrdinalIgnoreCase)){throw 'UserRoot must be on SourceRoot.'}

$roots=New-Object System.Collections.Generic.List[string]
$roots.Add($UserRoot)
foreach($r in $AdditionalRoots){
  if([string]::IsNullOrWhiteSpace($r)){continue}
  $nr=Normalize-Root $r
  if(-not $nr.StartsWith($SourceRoot,[System.StringComparison]::OrdinalIgnoreCase)){throw "AdditionalRoot must be on SourceRoot: $nr"}
  $roots.Add($nr)
}

$disk=Get-CheckedDisk $SourceRoot $ExpectedSerial
$DiskNumber=[int]$disk.Number
$DiskFriendlyName=[string]$disk.FriendlyName
$DiskReadOnlyAtStart=[bool]$disk.IsReadOnly
if($InterventionEndTime -lt $FirstDocumentedFailureTime){throw 'InterventionEndTime precedes FirstDocumentedFailureTime.'}

$exts=@('.docx','.xlsx','.xls','.doc','.pptx','.rtf','.txt','.md','.pdf','.psd','.ai','.indd','.jpg','.jpeg','.png','.tif','.tiff','.wav','.mp3','.m4a','.aac','.aup3','.csv','.json','.py','.zip','.rar')
$extSet=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach($e in $exts){[void]$extSet.Add($e)}

$outDir=Join-Path ([System.IO.Path]::GetDirectoryName($InventoryCsv)) ('file-audit-selector-v1_'+$RunID)
$tempDir=Join-Path $outDir '_temp'
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$planPath=Join-Path $outDir 'sample-plan.csv'
$summaryPath=Join-Path $outDir 'sample-summary.txt'
$oversizePath=Join-Path $outDir 'oversize-candidates.csv'
$interventionPath=Join-Path $outDir 'intervention-window.csv'
$afterPath=Join-Path $outDir 'after-intervention-review.csv'
$zeroPath=Join-Path $outDir 'zero-byte-candidates.csv'
$utf8=New-Object System.Text.UTF8Encoding($true)

$buckets=@('A_CRITICAL_DESKTOP','B_DESKTOP_4MONTHS','C_OTHER_4MONTHS','D_DESKTOP_HISTORICAL','E_OTHER_HISTORICAL','F_LEGACY_CONTROL')
$weights=@(1500.0,1200.0,700.0,300.0,100.0,100.0)
$budgets=@{}
$assigned=[UInt64]0
for($i=0;$i -lt $buckets.Count;$i++){
  if($i -eq $buckets.Count-1){$budgets[$buckets[$i]]=$MaxSampleBytes-$assigned}
  else{$v=[UInt64][math]::Floor($MaxSampleBytes*($weights[$i]/3900.0));$budgets[$buckets[$i]]=$v;$assigned+=$v}
}

$writers=@{}
foreach($b in $buckets){
  $w=New-Object System.IO.StreamWriter((Join-Path $tempDir ($b+'.csv')),$false,$utf8,262144)
  $w.WriteLine('FileID;SourceID;FullPath;RelativePath;Extension;SizeBytes;LastWriteTime;FileTypeGroup;TypeRank;TimeBand;TimeRank;IsDesktop;Bucket')
  $writers[$b]=$w
}

function New-ReviewWriter([string]$Path) {
  $w=New-Object System.IO.StreamWriter($Path,$false,$utf8,262144)
  $w.WriteLine('FileID;SourceID;FullPath;RelativePath;Extension;SizeBytes;LastWriteTime;FileTypeGroup;TimeBand;IsDesktop;Reason')
  return $w
}

$oversizeWriter=New-ReviewWriter $oversizePath
$interventionWriter=New-ReviewWriter $interventionPath
$afterWriter=New-ReviewWriter $afterPath
$zeroWriter=New-ReviewWriter $zeroPath

$parser=New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($InventoryCsv)
$parser.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(';')
$parser.HasFieldsEnclosedInQuotes=$true
$parser.TrimWhiteSpace=$false

[UInt64]$rowsRead=0
[UInt64]$relevant=0
[UInt64]$noiseSkipped=0
[UInt64]$outsideSkipped=0
[UInt64]$oversize=0
[UInt64]$zero=0
[UInt64]$intervention=0
[UInt64]$after=0
[UInt64]$sourceMismatch=0
$start=Get-Date
$lastUi=[datetime]::MinValue
$lastSafety=[datetime]::MinValue

try{
  $header=$parser.ReadFields()
  $idx=@{}
  for($i=0;$i -lt $header.Count;$i++){$idx[$header[$i]]=$i}
  foreach($name in @('FileID','SourceID','FullPath','RelativePath','Extension','SizeBytes','LastWriteTime')){
    if(-not $idx.ContainsKey($name)){throw "Required inventory column missing: $name"}
  }

  while(-not $parser.EndOfData){
    $row=$parser.ReadFields()
    $rowsRead++
    $now=Get-Date

    if(($now-$lastSafety).TotalSeconds -ge 30){[void](Assert-ReadOnly $DiskNumber $ExpectedSerial);$lastSafety=$now}
    if(($now-$lastUi).TotalMilliseconds -ge 750){
      $elapsed=$now-$start
      $rate=if($elapsed.TotalSeconds -gt 0){[double]$rowsRead/$elapsed.TotalSeconds}else{0}
      Write-Progress -Activity 'Pass 1/2 - Selecting audit candidates' -Status ('{0:N0} rows | {1:N0} relevant | {2:N0} rows/s' -f $rowsRead,$relevant,$rate)
      $lastUi=$now
    }

    if($row.Count -lt $header.Count){continue}
    $ext=$row[$idx['Extension']]
    if([string]::IsNullOrWhiteSpace($ext) -or -not $extSet.Contains($ext)){continue}
    $path=$row[$idx['FullPath']]
    if([string]::IsNullOrWhiteSpace($path)){continue}

    $allowed=$false
    foreach($r in $roots){if($path.StartsWith($r,[System.StringComparison]::OrdinalIgnoreCase)){$allowed=$true;break}}
    if(-not $allowed){$outsideSkipped++;continue}
    if(Is-Noise $path){$noiseSkipped++;continue}

    $sourceId=$row[$idx['SourceID']]
    if(-not [string]::IsNullOrWhiteSpace($sourceId)){
      if($sourceId -ne ('DISK-'+$ExpectedSerial)){$sourceMismatch++;continue}
    }

    [UInt64]$size=0
    if(-not [UInt64]::TryParse($row[$idx['SizeBytes']],[ref]$size)){continue}
    $stamp=[datetime]::MinValue
    if(-not [datetime]::TryParseExact($row[$idx['LastWriteTime']],'yyyy-MM-dd HH:mm:ss',[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$stamp)){continue}

    $desktop=($path -match '(?i)(^|\\)Desktop(\\|$)')
    $group=Get-Group $ext
    $band=Get-TimeBand $stamp
    $relevant++

    $common=@((Csv $row[$idx['FileID']]),(Csv $sourceId),(Csv $path),(Csv $row[$idx['RelativePath']]),(Csv $ext),([string]$size),(Csv $row[$idx['LastWriteTime']]),(Csv $group))

    if($band -eq 'INTERVENTION_WINDOW'){$intervention++;$interventionWriter.WriteLine((($common+@((Csv $band),(Csv ([string]$desktop)),(Csv 'Intervention window'))) -join ';'));continue}
    if($band -eq 'AFTER_INTERVENTION'){$after++;$afterWriter.WriteLine((($common+@((Csv $band),(Csv ([string]$desktop)),(Csv 'After intervention'))) -join ';'));continue}
    if($size -eq 0){$zero++;$zeroWriter.WriteLine((($common+@((Csv $band),(Csv ([string]$desktop)),(Csv 'Zero-byte'))) -join ';'));continue}
    if($size -gt $OversizeThresholdBytes){$oversize++;$oversizeWriter.WriteLine((($common+@((Csv $band),(Csv ([string]$desktop)),(Csv 'Oversize; manual review'))) -join ';'));continue}

    $bucket=Get-Bucket $desktop $band
    $line=@((Csv $row[$idx['FileID']]),(Csv $sourceId),(Csv $path),(Csv $row[$idx['RelativePath']]),(Csv $ext),([string]$size),(Csv $row[$idx['LastWriteTime']]),(Csv $group),([string](Get-TypeRank $ext)),(Csv $band),([string](Get-TimeRank $band)),(Csv ([string]$desktop)),(Csv $bucket)) -join ';'
    $writers[$bucket].WriteLine($line)
  }
}
finally{
  foreach($w in $writers.Values){$w.Flush();$w.Dispose()}
  foreach($w in @($oversizeWriter,$interventionWriter,$afterWriter,$zeroWriter)){$w.Flush();$w.Dispose()}
  $parser.Close()
  Write-Progress -Activity 'Pass 1/2 - Selecting audit candidates' -Completed
}

[void](Assert-ReadOnly $DiskNumber $ExpectedSerial)
if($sourceMismatch -gt 0){throw "SAFETY ABORT: $sourceMismatch inventory rows had an unexpected SourceID."}

$plan=New-Object System.IO.StreamWriter($planPath,$false,$utf8,1048576)
$plan.WriteLine('PlanNo;FileID;SourceID;OriginalPath;RelativePath;Extension;SizeBytes;LastWriteTime;FileTypeGroup;TimeBand;IsDesktop;Bucket;PlannedSampleRoot;SelectionReason')

[UInt64]$selectedBytes=0
[UInt64]$selectedFiles=0
[UInt64]$carry=0
$byBucket=@{}
$byExtension=@{}

try{
  foreach($bucket in $buckets){
    [UInt64]$available=[UInt64]$budgets[$bucket]+$carry
    $rows=Import-Csv -LiteralPath (Join-Path $tempDir ($bucket+'.csv')) -Delimiter ';' -Encoding UTF8
    $ordered=$rows | Sort-Object @{Expression={[int]$_.TimeRank};Ascending=$true},@{Expression={[int]$_.TypeRank};Ascending=$true},@{Expression={[datetime]$_.LastWriteTime};Descending=$true},@{Expression={[UInt64]$_.SizeBytes};Ascending=$true}
    [UInt64]$bucketFiles=0
    [UInt64]$bucketBytes=0

    foreach($r in $ordered){
      [UInt64]$size=[UInt64]$r.SizeBytes
      if(($selectedBytes+$size) -gt $MaxSampleBytes -or $size -gt $available){continue}
      $selectedFiles++;$selectedBytes+=$size;$available-=$size;$bucketFiles++;$bucketBytes+=$size
      $ek=$r.Extension.ToLowerInvariant()
      if(-not $byExtension.ContainsKey($ek)){$byExtension[$ek]=[UInt64]0}
      $byExtension[$ek]++
      $reason=$r.TimeBand+'; '+$r.FileTypeGroup
      if($r.IsDesktop -eq 'True'){$reason+='; Desktop'}
      $line=@(([string]$selectedFiles),(Csv $r.FileID),(Csv $r.SourceID),(Csv $r.FullPath),(Csv $r.RelativePath),(Csv $r.Extension),([string]$size),(Csv $r.LastWriteTime),(Csv $r.FileTypeGroup),(Csv $r.TimeBand),(Csv $r.IsDesktop),(Csv $bucket),(Csv $PlannedSampleFolderName),(Csv $reason)) -join ';'
      $plan.WriteLine($line)
    }

    $byBucket[$bucket]=@($bucketFiles,$bucketBytes)
    $carry=$available
    $pos=[array]::IndexOf($buckets,$bucket)+1
    Write-Progress -Activity 'Pass 2/2 - Building sample plan' -Status ('{0:N0} files | {1:N3} decimal GB' -f $selectedFiles,($selectedBytes/1000000000)) -CurrentOperation $bucket -PercentComplete (($pos/[double]$buckets.Count)*100)
  }
}
finally{
  $plan.Flush();$plan.Dispose()
  Write-Progress -Activity 'Pass 2/2 - Building sample plan' -Completed
}

$diskEnd=Assert-ReadOnly $DiskNumber $ExpectedSerial

$summary=New-Object System.Collections.Generic.List[string]
$summary.Add('FILE AUDIT SELECTOR V1')
$summary.Add('======================')
$summary.Add("Version=$ScriptVersion")
$summary.Add("InventoryCsv=$InventoryCsv")
$summary.Add("SourceRoot=$SourceRoot")
$summary.Add("DiskFriendlyName=$DiskFriendlyName")
$summary.Add("DiskReadOnlyAtStart=$DiskReadOnlyAtStart")
$summary.Add("DiskReadOnlyAtEnd=$($diskEnd.IsReadOnly)")
$summary.Add("UserRoot=$UserRoot")
$summary.Add("FirstDocumentedFailureTime=$($FirstDocumentedFailureTime.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("InterventionEndTime=$($InterventionEndTime.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("HistoricalStartYear=$HistoricalStartYear")
$summary.Add("MaxSampleBytes=$MaxSampleBytes")
$summary.Add("PlannedSampleFolder=$PlannedSampleFolderName")
$summary.Add("RowsRead=$rowsRead")
$summary.Add("RelevantRows=$relevant")
$summary.Add("NoiseRowsSkipped=$noiseSkipped")
$summary.Add("OutsideRootRowsSkipped=$outsideSkipped")
$summary.Add("InterventionRows=$intervention")
$summary.Add("AfterInterventionRows=$after")
$summary.Add("OversizeRows=$oversize")
$summary.Add("ZeroByteRows=$zero")
$summary.Add("SelectedFiles=$selectedFiles")
$summary.Add("SelectedBytes=$selectedBytes")
$summary.Add("SelectedDecimalGB=$([math]::Round($selectedBytes/1000000000,3))")
$summary.Add("RemainingBudgetBytes=$($MaxSampleBytes-$selectedBytes)")
foreach($b in $buckets){$summary.Add("$b.Files=$($byBucket[$b][0])");$summary.Add("$b.Bytes=$($byBucket[$b][1])")}
foreach($e in ($byExtension.Keys|Sort-Object)){$summary.Add("$e=$($byExtension[$e])")}
$summary.Add("COPY_PERFORMED=False")
$summary | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host
Write-Host 'FILE AUDIT SELECTOR V1 - DONE'
Write-Host ("Rows read           : {0:N0}" -f $rowsRead)
Write-Host ("Relevant candidates : {0:N0}" -f $relevant)
Write-Host ("Selected files      : {0:N0}" -f $selectedFiles)
Write-Host ("Selected decimal GB : {0:N3}" -f ($selectedBytes/1000000000))
Write-Host ("Source read-only    : {0}" -f $diskEnd.IsReadOnly)
Write-Host "Plan                 : $planPath"
Write-Host "Summary              : $summaryPath"
Write-Host "Oversize             : $oversizePath"
Write-Host "Intervention         : $interventionPath"
Write-Host "Zero-byte            : $zeroPath"
Write-Host 'No source files were copied.'
