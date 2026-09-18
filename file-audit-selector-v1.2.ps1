#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
  Builds a stratified, size-limited recovery audit sample plan.

.DESCRIPTION
  Version 1.2 is planning-only. It reads the file-inventator master CSV and
  creates a deterministic sample plan without opening or copying source files.

  Safety:
    - the mounted source disk serial must match ExpectedSerial;
    - Windows must report the source disk as read-only;
    - read-only state is checked before the run, every 30 seconds while the
      inventory is scanned, during plan construction, and again at the end;
    - inventory and output must be on a volume other than the source.

  Time strata are mutually exclusive and measured backwards from
  FirstDocumentedFailureTime:
    T01  <= 1 day
    T02  >1 day to 3 days
    T03  >3 days to 1 week
    T04  >1 week to 2 weeks
    T05  >2 weeks to 3 weeks
    T06  >3 weeks to 1 calendar month
    T07  >1 month to 6 weeks
    T08  >6 weeks to 2 calendar months
    T09  >2 to 3 calendar months
    T10  >3 to 6 calendar months
    T11  >6 months to 1 year
    T12  >1 to 2 years
    T13  >2 to 3 years
    T14  >3 to 5 years

  Files older than five years are listed separately and are not automatically
  selected. Files inside the intervention window and after it are also listed
  separately.

  Type priorities within every time stratum:
    1 CREATIVE_NATIVE  .psd .ai .indd .aup3
    2 OFFICE_AUTHORING .docx .xlsx .xls .doc .pptx .rtf
    3 MEDIA            .jpg .jpeg .png .tif .tiff .wav .mp3 .m4a .aac
    4 PDF              .pdf
    5 DATA_CODE        .csv .json .py
    6 TEXT_MARKUP      .txt .md
    7 ARCHIVE          .zip .rar

  The selector uses two simultaneous stratifications:
    - every time stratum receives a nominal share of the total byte budget;
    - every type group receives a global target share of the total byte budget.

  Phase A reserves the time-by-type matrix so all 14 time strata are
  represented. Phase B tops up global type targets in priority order, with
  creative native files first. Phase C uses any remaining bytes in type
  priority order. Desktop files are preferred, then newer files.

  Files above OversizeThresholdBytes are listed separately for manual review.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File ".\file-audit-selector-v1.2.ps1" -InventoryCsv ".\file-inventator-v1_master.csv" -SourceRoot "X:\" -ExpectedSerial "SERIAL_NUMBER" -UserRoot "X:\Users\PROFILE\" -FirstDocumentedFailureTime "2026-07-26 20:32:24" -InterventionEndTime "2026-07-27 04:00:00" -PlannedSampleFolderName "!AnalyseSampleRecovery"
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
  [Parameter()][string]$PlannedSampleFolderName='!AnalyseSampleRecovery',
  [Parameter()][switch]$IncludeNoisePaths
)

$ErrorActionPreference='Stop'
$ScriptVersion='1.2'
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

function Get-TypeGroup([string]$Ext) {
  switch($Ext.ToLowerInvariant()){
    {$_ -in @('.psd','.ai','.indd','.aup3')}{return 'CREATIVE_NATIVE'}
    {$_ -in @('.docx','.xlsx','.xls','.doc','.pptx','.rtf')}{return 'OFFICE_AUTHORING'}
    {$_ -in @('.jpg','.jpeg','.png','.tif','.tiff','.wav','.mp3','.m4a','.aac')}{return 'MEDIA'}
    '.pdf'{return 'PDF'}
    {$_ -in @('.csv','.json','.py')}{return 'DATA_CODE'}
    {$_ -in @('.txt','.md')}{return 'TEXT_MARKUP'}
    {$_ -in @('.zip','.rar')}{return 'ARCHIVE'}
    default{return 'OTHER'}
  }
}

function Get-TypeRank([string]$Group) {
  switch($Group){
    'CREATIVE_NATIVE'{return 1}
    'OFFICE_AUTHORING'{return 2}
    'MEDIA'{return 3}
    'PDF'{return 4}
    'DATA_CODE'{return 5}
    'TEXT_MARKUP'{return 6}
    'ARCHIVE'{return 7}
    default{return 99}
  }
}

function Get-TimeBand([datetime]$Stamp) {
  $a=$script:FirstDocumentedFailureTime
  if($Stamp -ge $a -and $Stamp -le $script:InterventionEndTime){return 'INTERVENTION_WINDOW'}
  if($Stamp -gt $script:InterventionEndTime){return 'AFTER_INTERVENTION'}
  if($Stamp -ge $a.AddDays(-1)){return 'T01_1_DAY'}
  if($Stamp -ge $a.AddDays(-3)){return 'T02_3_DAYS'}
  if($Stamp -ge $a.AddDays(-7)){return 'T03_1_WEEK'}
  if($Stamp -ge $a.AddDays(-14)){return 'T04_2_WEEKS'}
  if($Stamp -ge $a.AddDays(-21)){return 'T05_3_WEEKS'}
  if($Stamp -ge $a.AddMonths(-1)){return 'T06_1_MONTH'}
  if($Stamp -ge $a.AddDays(-42)){return 'T07_6_WEEKS'}
  if($Stamp -ge $a.AddMonths(-2)){return 'T08_2_MONTHS'}
  if($Stamp -ge $a.AddMonths(-3)){return 'T09_3_MONTHS'}
  if($Stamp -ge $a.AddMonths(-6)){return 'T10_6_MONTHS'}
  if($Stamp -ge $a.AddYears(-1)){return 'T11_1_YEAR'}
  if($Stamp -ge $a.AddYears(-2)){return 'T12_2_YEARS'}
  if($Stamp -ge $a.AddYears(-3)){return 'T13_3_YEARS'}
  if($Stamp -ge $a.AddYears(-5)){return 'T14_5_YEARS'}
  return 'OLDER_THAN_5_YEARS'
}

function Is-Noise([string]$Path) {
  if($script:IncludeNoisePaths){return $false}
  foreach($p in @('\AppData\','\node_modules\','\site-packages\','\__pycache__\','\.git\','\.cache\','\venv\','\.venv\','\OneDriveTemp\')){
    if($Path.IndexOf($p,[System.StringComparison]::OrdinalIgnoreCase) -ge 0){return $true}
  }
  return $false
}

function New-ReviewWriter([string]$Path,[System.Text.Encoding]$Encoding) {
  $w=New-Object System.IO.StreamWriter($Path,$false,$Encoding,262144)
  $w.WriteLine('FileID;SourceID;FullPath;RelativePath;Extension;SizeBytes;LastWriteTime;TypeGroup;TimeBand;IsDesktop;Reason')
  return $w
}

function Select-RowsWithinBudget {
  param(
    [object[]]$Rows,
    [UInt64]$Budget,
    [System.Collections.Generic.HashSet[string]]$SelectedIds,
    [string]$Stage,
    [string]$Band,
    [System.IO.StreamWriter]$PlanWriter
  )

  [UInt64]$used=0
  [UInt64]$count=0

  foreach($r in $Rows){
    if($SelectedIds.Contains($r.FileID)){continue}
    [UInt64]$size=[UInt64]$r.SizeBytes
    [UInt64]$remaining=if($Budget -gt $used){$Budget-$used}else{0}
    if($size -gt $remaining){continue}
    if(($script:SelectedBytes+$size) -gt $script:MaxSampleBytes){continue}

    [void]$SelectedIds.Add($r.FileID)
    $script:SelectedFiles++
    $script:SelectedBytes+=$size
    $used+=$size
    $count++

    if(-not $script:SelectedByTimeBand.ContainsKey($Band)){
      $script:SelectedByTimeBand[$Band]=[UInt64]0
      $script:SelectedBytesByTimeBand[$Band]=[UInt64]0
    }
    $script:SelectedByTimeBand[$Band]++
    $script:SelectedBytesByTimeBand[$Band]+=$size

    if(-not $script:SelectedByTypeGroup.ContainsKey($r.TypeGroup)){
      $script:SelectedByTypeGroup[$r.TypeGroup]=[UInt64]0
      $script:SelectedBytesByTypeGroup[$r.TypeGroup]=[UInt64]0
    }
    $script:SelectedByTypeGroup[$r.TypeGroup]++
    $script:SelectedBytesByTypeGroup[$r.TypeGroup]+=$size

    $extKey=$r.Extension.ToLowerInvariant()
    if(-not $script:SelectedByExtension.ContainsKey($extKey)){
      $script:SelectedByExtension[$extKey]=[UInt64]0
      $script:SelectedBytesByExtension[$extKey]=[UInt64]0
    }
    $script:SelectedByExtension[$extKey]++
    $script:SelectedBytesByExtension[$extKey]+=$size

    if($r.IsDesktop -eq 'True'){
      $script:SelectedDesktopFiles++
      $script:SelectedDesktopBytes+=$size
    }else{
      $script:SelectedOtherFiles++
      $script:SelectedOtherBytes+=$size
    }

    $reason=$Band+'; '+$r.TypeGroup
    if($r.IsDesktop -eq 'True'){$reason+='; Desktop'}

    $line=@(
      ([string]$script:SelectedFiles),
      (Csv $r.FileID),
      (Csv $r.SourceID),
      (Csv $r.FullPath),
      (Csv $r.RelativePath),
      (Csv $r.Extension),
      ([string]$size),
      (Csv $r.LastWriteTime),
      (Csv $r.TypeGroup),
      (Csv $Band),
      (Csv $r.IsDesktop),
      (Csv $Stage),
      (Csv $script:PlannedSampleFolderName),
      (Csv $reason)
    ) -join ';'

    $PlanWriter.WriteLine($line)
    if($used -ge $Budget){break}
  }

  return [PSCustomObject]@{Count=$count;Bytes=$used}
}

$InventoryCsv=[System.IO.Path]::GetFullPath($InventoryCsv)
if(-not(Test-Path -LiteralPath $InventoryCsv)){throw "Inventory CSV not found: $InventoryCsv"}

$SourceRoot=Normalize-Root $SourceRoot
$UserRoot=Normalize-Root $UserRoot
$ExpectedSerial=$ExpectedSerial.Trim()

$inventoryDrive=[System.IO.Path]::GetPathRoot($InventoryCsv)
if($inventoryDrive.TrimEnd('\').Equals($SourceRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){throw 'SAFETY ABORT: inventory/output location must not be on the source volume.'}
if(-not $UserRoot.StartsWith($SourceRoot,[System.StringComparison]::OrdinalIgnoreCase)){throw 'UserRoot must be on SourceRoot.'}

$roots=New-Object System.Collections.Generic.List[string]
$roots.Add($UserRoot)
foreach($r in $AdditionalRoots){
  if([string]::IsNullOrWhiteSpace($r)){continue}
  $nr=Normalize-Root $r
  if(-not $nr.StartsWith($SourceRoot,[System.StringComparison]::OrdinalIgnoreCase)){throw "AdditionalRoot must be on SourceRoot: $nr"}
  $roots.Add($nr)
}

if($InterventionEndTime -lt $FirstDocumentedFailureTime){throw 'InterventionEndTime precedes FirstDocumentedFailureTime.'}

$disk=Get-CheckedDisk $SourceRoot $ExpectedSerial
$DiskNumber=[int]$disk.Number
$DiskFriendlyName=[string]$disk.FriendlyName
$DiskReadOnlyAtStart=[bool]$disk.IsReadOnly

$RelevantExtensions=@('.docx','.xlsx','.xls','.doc','.pptx','.rtf','.txt','.md','.pdf','.psd','.ai','.indd','.jpg','.jpeg','.png','.tif','.tiff','.wav','.mp3','.m4a','.aac','.aup3','.csv','.json','.py','.zip','.rar')
$ExtSet=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach($e in $RelevantExtensions){[void]$ExtSet.Add($e)}

$TimeBands=@('T01_1_DAY','T02_3_DAYS','T03_1_WEEK','T04_2_WEEKS','T05_3_WEEKS','T06_1_MONTH','T07_6_WEEKS','T08_2_MONTHS','T09_3_MONTHS','T10_6_MONTHS','T11_1_YEAR','T12_2_YEARS','T13_3_YEARS','T14_5_YEARS')

$TimeRankMap=@{}
for($i=0;$i -lt $TimeBands.Count;$i++){
  $TimeRankMap[$TimeBands[$i]]=$i+1
}

$TimeWeights=[ordered]@{
  T01_1_DAY=13
  T02_3_DAYS=10
  T03_1_WEEK=8
  T04_2_WEEKS=7
  T05_3_WEEKS=6
  T06_1_MONTH=6
  T07_6_WEEKS=5
  T08_2_MONTHS=5
  T09_3_MONTHS=5
  T10_6_MONTHS=8
  T11_1_YEAR=8
  T12_2_YEARS=7
  T13_3_YEARS=6
  T14_5_YEARS=6
}

$TypeWeights=[ordered]@{
  CREATIVE_NATIVE=35
  OFFICE_AUTHORING=20
  MEDIA=20
  PDF=12
  DATA_CODE=7
  TEXT_MARKUP=5
  ARCHIVE=1
}

$outDir=Join-Path ([System.IO.Path]::GetDirectoryName($InventoryCsv)) ('file-audit-selector-v1.2_'+$RunID)
$tempDir=Join-Path $outDir '_temp'
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$candidatePath=Join-Path $tempDir 'candidates.csv'
$planPath=Join-Path $outDir 'sample-plan.csv'
$summaryPath=Join-Path $outDir 'sample-summary.txt'
$oversizePath=Join-Path $outDir 'oversize-candidates.csv'
$interventionPath=Join-Path $outDir 'intervention-window.csv'
$afterPath=Join-Path $outDir 'after-intervention-review.csv'
$olderPath=Join-Path $outDir 'older-than-5-years-review.csv'
$zeroPath=Join-Path $outDir 'zero-byte-candidates.csv'

$utf8=New-Object System.Text.UTF8Encoding($true)

$candidateWriter=New-Object System.IO.StreamWriter($candidatePath,$false,$utf8,1048576)
$candidateWriter.WriteLine('FileID;SourceID;FullPath;RelativePath;Extension;SizeBytes;LastWriteTime;TypeGroup;TypeRank;TimeBand;TimeRank;IsDesktop')
$oversizeWriter=New-ReviewWriter $oversizePath $utf8
$interventionWriter=New-ReviewWriter $interventionPath $utf8
$afterWriter=New-ReviewWriter $afterPath $utf8
$olderWriter=New-ReviewWriter $olderPath $utf8
$zeroWriter=New-ReviewWriter $zeroPath $utf8

$parser=New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($InventoryCsv)
$parser.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(';')
$parser.HasFieldsEnclosedInQuotes=$true
$parser.TrimWhiteSpace=$false

[UInt64]$RowsRead=0
[UInt64]$RelevantRows=0
[UInt64]$NoiseRowsSkipped=0
[UInt64]$OutsideRootRowsSkipped=0
[UInt64]$OversizeRows=0
[UInt64]$ZeroByteRows=0
[UInt64]$InterventionRows=0
[UInt64]$AfterInterventionRows=0
[UInt64]$OlderThanFiveYearsRows=0
[UInt64]$SourceMismatchRows=0

$CandidateFilesByTimeBand=@{}
$CandidateBytesByTimeBand=@{}
$CandidateFilesByTypeGroup=@{}
$CandidateBytesByTypeGroup=@{}

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
    $RowsRead++
    $now=Get-Date

    if(($now-$lastSafety).TotalSeconds -ge 30){[void](Assert-ReadOnly $DiskNumber $ExpectedSerial);$lastSafety=$now}

    if(($now-$lastUi).TotalMilliseconds -ge 750){
      $elapsed=$now-$start
      $rate=if($elapsed.TotalSeconds -gt 0){[double]$RowsRead/$elapsed.TotalSeconds}else{0}
      Write-Progress -Activity 'Pass 1/2 - Classifying recovery candidates' -Status ('{0:N0} rows | {1:N0} relevant | {2:N0} rows/s' -f $RowsRead,$RelevantRows,$rate)
      $lastUi=$now
    }

    if($row.Count -lt $header.Count){continue}
    $ext=$row[$idx['Extension']]
    if([string]::IsNullOrWhiteSpace($ext) -or -not $ExtSet.Contains($ext)){continue}

    $path=$row[$idx['FullPath']]
    if([string]::IsNullOrWhiteSpace($path)){continue}

    $allowed=$false
    foreach($r in $roots){
      if($path.StartsWith($r,[System.StringComparison]::OrdinalIgnoreCase)){$allowed=$true;break}
    }
    if(-not $allowed){$OutsideRootRowsSkipped++;continue}
    if(Is-Noise $path){$NoiseRowsSkipped++;continue}

    $sourceId=$row[$idx['SourceID']]
    if(-not [string]::IsNullOrWhiteSpace($sourceId)){
      if($sourceId -ne ('DISK-'+$ExpectedSerial)){$SourceMismatchRows++;continue}
    }

    [UInt64]$size=0
    if(-not [UInt64]::TryParse($row[$idx['SizeBytes']],[ref]$size)){continue}

    $stamp=[datetime]::MinValue
    if(-not [datetime]::TryParseExact($row[$idx['LastWriteTime']],'yyyy-MM-dd HH:mm:ss',[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$stamp)){continue}

    $desktop=($path -match '(?i)(^|\\)Desktop(\\|$)')
    $typeGroup=Get-TypeGroup $ext
    $typeRank=Get-TypeRank $typeGroup
    $timeBand=Get-TimeBand $stamp
    $RelevantRows++

    $common=@((Csv $row[$idx['FileID']]),(Csv $sourceId),(Csv $path),(Csv $row[$idx['RelativePath']]),(Csv $ext),([string]$size),(Csv $row[$idx['LastWriteTime']]),(Csv $typeGroup),(Csv $timeBand),(Csv ([string]$desktop)))

    if($timeBand -eq 'INTERVENTION_WINDOW'){$InterventionRows++;$interventionWriter.WriteLine((($common+(Csv 'Intervention window')) -join ';'));continue}
    if($timeBand -eq 'AFTER_INTERVENTION'){$AfterInterventionRows++;$afterWriter.WriteLine((($common+(Csv 'After intervention')) -join ';'));continue}
    if($timeBand -eq 'OLDER_THAN_5_YEARS'){$OlderThanFiveYearsRows++;$olderWriter.WriteLine((($common+(Csv 'Older than five years; review only')) -join ';'));continue}
    if($size -eq 0){$ZeroByteRows++;$zeroWriter.WriteLine((($common+(Csv 'Zero-byte candidate')) -join ';'));continue}
    if($size -gt $OversizeThresholdBytes){$OversizeRows++;$oversizeWriter.WriteLine((($common+(Csv 'Oversize; manual review')) -join ';'));continue}

    if(-not $CandidateFilesByTimeBand.ContainsKey($timeBand)){
      $CandidateFilesByTimeBand[$timeBand]=[UInt64]0
      $CandidateBytesByTimeBand[$timeBand]=[UInt64]0
    }
    if(-not $CandidateFilesByTypeGroup.ContainsKey($typeGroup)){
      $CandidateFilesByTypeGroup[$typeGroup]=[UInt64]0
      $CandidateBytesByTypeGroup[$typeGroup]=[UInt64]0
    }

    $CandidateFilesByTimeBand[$timeBand]++
    $CandidateBytesByTimeBand[$timeBand]+=$size
    $CandidateFilesByTypeGroup[$typeGroup]++
    $CandidateBytesByTypeGroup[$typeGroup]+=$size

    $candidateWriter.WriteLine((@((Csv $row[$idx['FileID']]),(Csv $sourceId),(Csv $path),(Csv $row[$idx['RelativePath']]),(Csv $ext),([string]$size),(Csv $row[$idx['LastWriteTime']]),(Csv $typeGroup),([string]$typeRank),(Csv $timeBand),([string]$TimeRankMap[$timeBand]),(Csv ([string]$desktop))) -join ';'))
  }
}
finally{
  $candidateWriter.Flush()
  $candidateWriter.Dispose()
  foreach($w in @($oversizeWriter,$interventionWriter,$afterWriter,$olderWriter,$zeroWriter)){$w.Flush();$w.Dispose()}
  $parser.Close()
  Write-Progress -Activity 'Pass 1/2 - Classifying recovery candidates' -Completed
}

[void](Assert-ReadOnly $DiskNumber $ExpectedSerial)
if($SourceMismatchRows -gt 0){throw "SAFETY ABORT: $SourceMismatchRows inventory rows had an unexpected SourceID."}

$candidates=Import-Csv -LiteralPath $candidatePath -Delimiter ';' -Encoding UTF8
$planWriter=New-Object System.IO.StreamWriter($planPath,$false,$utf8,1048576)
$planWriter.WriteLine('PlanNo;FileID;SourceID;OriginalPath;RelativePath;Extension;SizeBytes;LastWriteTime;TypeGroup;TimeBand;IsDesktop;SelectionStage;PlannedSampleRoot;SelectionReason')

$SelectedIds=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

$script:SelectedFiles=[UInt64]0
$script:SelectedBytes=[UInt64]0
$script:SelectedByTimeBand=@{}
$script:SelectedBytesByTimeBand=@{}
$script:SelectedByTypeGroup=@{}
$script:SelectedBytesByTypeGroup=@{}
$script:SelectedByExtension=@{}
$script:SelectedBytesByExtension=@{}
$script:SelectedDesktopFiles=[UInt64]0
$script:SelectedDesktopBytes=[UInt64]0
$script:SelectedOtherFiles=[UInt64]0
$script:SelectedOtherBytes=[UInt64]0

$TimeBudgetBytes=@{}
[UInt64]$assignedTime=0
for($i=0;$i -lt $TimeBands.Count;$i++){
  $band=$TimeBands[$i]
  if($i -eq ($TimeBands.Count-1)){
    $TimeBudgetBytes[$band]=$MaxSampleBytes-$assignedTime
  }else{
    [UInt64]$v=[UInt64][math]::Floor($MaxSampleBytes*([double]$TimeWeights[$band]/100.0))
    $TimeBudgetBytes[$band]=$v
    $assignedTime+=$v
  }
}

$GlobalTypeTargetBytes=@{}
[UInt64]$assignedType=0
$typeKeys=@($TypeWeights.Keys)
for($i=0;$i -lt $typeKeys.Count;$i++){
  $group=$typeKeys[$i]
  if($i -eq ($typeKeys.Count-1)){
    $GlobalTypeTargetBytes[$group]=$MaxSampleBytes-$assignedType
  }else{
    [UInt64]$v=[UInt64][math]::Floor($MaxSampleBytes*([double]$TypeWeights[$group]/100.0))
    $GlobalTypeTargetBytes[$group]=$v
    $assignedType+=$v
  }
}

try{
  # Phase A: reserve every time-by-type cell.
  for($i=0;$i -lt $TimeBands.Count;$i++){
    [void](Assert-ReadOnly $DiskNumber $ExpectedSerial)
    $band=$TimeBands[$i]
    [UInt64]$bandBudget=[UInt64]$TimeBudgetBytes[$band]

    foreach($typeGroup in $TypeWeights.Keys){
      [UInt64]$cellBudget=[UInt64][math]::Floor($bandBudget*([double]$TypeWeights[$typeGroup]/100.0))
      if($cellBudget -eq 0){continue}

      $typed=@(
        $candidates |
        Where-Object {$_.TimeBand -eq $band -and $_.TypeGroup -eq $typeGroup} |
        Sort-Object @{Expression={if($_.IsDesktop -eq 'True'){0}else{1}};Ascending=$true},@{Expression={[datetime]$_.LastWriteTime};Descending=$true},@{Expression={$_.FileID};Ascending=$true}
      )

      if($typed.Count -gt 0){
        [void](Select-RowsWithinBudget -Rows $typed -Budget $cellBudget -SelectedIds $SelectedIds -Stage 'MATRIX_RESERVED' -Band $band -PlanWriter $planWriter)
      }
    }

    Write-Progress -Activity 'Pass 2/2 - Building stratified sample plan' -Status ('Phase A matrix | {0:N0} files | {1:N3} decimal GB' -f $script:SelectedFiles,($script:SelectedBytes/1000000000)) -CurrentOperation $band -PercentComplete ((($i+1)/[double]$TimeBands.Count)*45)
  }

  # Phase B: enforce global type targets. Creative-native is topped up first.
  $typeIndex=0
  foreach($typeGroup in $TypeWeights.Keys){
    [void](Assert-ReadOnly $DiskNumber $ExpectedSerial)
    [UInt64]$already=if($script:SelectedBytesByTypeGroup.ContainsKey($typeGroup)){$script:SelectedBytesByTypeGroup[$typeGroup]}else{0}
    [UInt64]$target=[UInt64]$GlobalTypeTargetBytes[$typeGroup]
    [UInt64]$gap=if($target -gt $already){$target-$already}else{0}

    if($gap -gt 0){
      $typed=@(
        $candidates |
        Where-Object {$_.TypeGroup -eq $typeGroup} |
        Sort-Object @{Expression={[int]$_.TimeRank};Ascending=$true},@{Expression={if($_.IsDesktop -eq 'True'){0}else{1}};Ascending=$true},@{Expression={[datetime]$_.LastWriteTime};Descending=$true},@{Expression={$_.FileID};Ascending=$true}
      )

      if($typed.Count -gt 0){
        [void](Select-RowsWithinBudget -Rows $typed -Budget $gap -SelectedIds $SelectedIds -Stage 'GLOBAL_TYPE_TOPUP' -Band 'MIXED_TIME_BANDS' -PlanWriter $planWriter)
      }
    }

    $typeIndex++
    Write-Progress -Activity 'Pass 2/2 - Building stratified sample plan' -Status ('Phase B global targets | {0:N0} files | {1:N3} decimal GB' -f $script:SelectedFiles,($script:SelectedBytes/1000000000)) -CurrentOperation $typeGroup -PercentComplete (45+(($typeIndex/[double]$TypeWeights.Count)*40))
  }

  # Phase C: use any remaining bytes in creative-first priority order.
  [UInt64]$remainingTotal=if($MaxSampleBytes -gt $script:SelectedBytes){$MaxSampleBytes-$script:SelectedBytes}else{0}
  if($remainingTotal -gt 0){
    $fillRows=@(
      $candidates |
      Sort-Object @{Expression={[int]$_.TypeRank};Ascending=$true},@{Expression={[int]$_.TimeRank};Ascending=$true},@{Expression={if($_.IsDesktop -eq 'True'){0}else{1}};Ascending=$true},@{Expression={[datetime]$_.LastWriteTime};Descending=$true},@{Expression={$_.FileID};Ascending=$true}
    )

    [void](Select-RowsWithinBudget -Rows $fillRows -Budget $remainingTotal -SelectedIds $SelectedIds -Stage 'FINAL_PRIORITY_FILL' -Band 'MIXED_TIME_BANDS' -PlanWriter $planWriter)
  }

  Write-Progress -Activity 'Pass 2/2 - Building stratified sample plan' -Status ('Phase C final fill | {0:N0} files | {1:N3} decimal GB' -f $script:SelectedFiles,($script:SelectedBytes/1000000000)) -PercentComplete 100
}
finally{
  $planWriter.Flush()
  $planWriter.Dispose()
  Write-Progress -Activity 'Pass 2/2 - Building stratified sample plan' -Completed
}

$diskEnd=Assert-ReadOnly $DiskNumber $ExpectedSerial

$summary=New-Object System.Collections.Generic.List[string]
$summary.Add('FILE AUDIT SELECTOR V1.2')
$summary.Add('========================')
$summary.Add("Version=$ScriptVersion")
$summary.Add("InventoryCsv=$InventoryCsv")
$summary.Add("SourceRoot=$SourceRoot")
$summary.Add("DiskFriendlyName=$DiskFriendlyName")
$summary.Add("DiskReadOnlyAtStart=$DiskReadOnlyAtStart")
$summary.Add("DiskReadOnlyAtEnd=$($diskEnd.IsReadOnly)")
$summary.Add("UserRoot=$UserRoot")
$summary.Add("FirstDocumentedFailureTime=$($FirstDocumentedFailureTime.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("InterventionEndTime=$($InterventionEndTime.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("MaxSampleBytes=$MaxSampleBytes")
$summary.Add("OversizeThresholdBytes=$OversizeThresholdBytes")
$summary.Add("PlannedSampleFolder=$PlannedSampleFolderName")
$summary.Add('')
$summary.Add("RowsRead=$RowsRead")
$summary.Add("RelevantRows=$RelevantRows")
$summary.Add("NoiseRowsSkipped=$NoiseRowsSkipped")
$summary.Add("OutsideRootRowsSkipped=$OutsideRootRowsSkipped")
$summary.Add("InterventionRows=$InterventionRows")
$summary.Add("AfterInterventionRows=$AfterInterventionRows")
$summary.Add("OlderThanFiveYearsRows=$OlderThanFiveYearsRows")
$summary.Add("OversizeRows=$OversizeRows")
$summary.Add("ZeroByteRows=$ZeroByteRows")
$summary.Add('')
$summary.Add("SelectedFiles=$($script:SelectedFiles)")
$summary.Add("SelectedBytes=$($script:SelectedBytes)")
$summary.Add("SelectedDecimalGB=$([math]::Round($script:SelectedBytes/1000000000,3))")
$summary.Add("RemainingBudgetBytes=$($MaxSampleBytes-$script:SelectedBytes)")
$summary.Add("SelectedDesktopFiles=$($script:SelectedDesktopFiles)")
$summary.Add("SelectedDesktopBytes=$($script:SelectedDesktopBytes)")
$summary.Add("SelectedOtherFiles=$($script:SelectedOtherFiles)")
$summary.Add("SelectedOtherBytes=$($script:SelectedOtherBytes)")
$summary.Add('')
$summary.Add('TIME BAND BUDGETS AND RESULTS')

foreach($band in $TimeBands){
  $candidateFiles=if($CandidateFilesByTimeBand.ContainsKey($band)){$CandidateFilesByTimeBand[$band]}else{0}
  $candidateBytes=if($CandidateBytesByTimeBand.ContainsKey($band)){$CandidateBytesByTimeBand[$band]}else{0}
  $selectedFiles=if($script:SelectedByTimeBand.ContainsKey($band)){$script:SelectedByTimeBand[$band]}else{0}
  $selectedBytes=if($script:SelectedBytesByTimeBand.ContainsKey($band)){$script:SelectedBytesByTimeBand[$band]}else{0}

  $summary.Add("$band.WeightPercent=$($TimeWeights[$band])")
  $summary.Add("$band.NominalBudgetBytes=$($TimeBudgetBytes[$band])")
  $summary.Add("$band.CandidateFiles=$candidateFiles")
  $summary.Add("$band.CandidateBytes=$candidateBytes")
  $summary.Add("$band.SelectedFiles=$selectedFiles")
  $summary.Add("$band.SelectedBytes=$selectedBytes")
}

$summary.Add('')
$summary.Add('GLOBAL TYPE TARGETS AND RESULTS')

foreach($group in $TypeWeights.Keys){
  $candidateFiles=if($CandidateFilesByTypeGroup.ContainsKey($group)){$CandidateFilesByTypeGroup[$group]}else{0}
  $candidateBytes=if($CandidateBytesByTypeGroup.ContainsKey($group)){$CandidateBytesByTypeGroup[$group]}else{0}
  $selectedFiles=if($script:SelectedByTypeGroup.ContainsKey($group)){$script:SelectedByTypeGroup[$group]}else{0}
  $selectedBytes=if($script:SelectedBytesByTypeGroup.ContainsKey($group)){$script:SelectedBytesByTypeGroup[$group]}else{0}

  $actualPercent=if($MaxSampleBytes -gt 0){[math]::Round(([double]$selectedBytes/[double]$MaxSampleBytes)*100,2)}else{0}
  $summary.Add("$group.TargetPercent=$($TypeWeights[$group])")
  $summary.Add("$group.TargetBytes=$($GlobalTypeTargetBytes[$group])")
  $summary.Add("$group.CandidateFiles=$candidateFiles")
  $summary.Add("$group.CandidateBytes=$candidateBytes")
  $summary.Add("$group.SelectedFiles=$selectedFiles")
  $summary.Add("$group.SelectedBytes=$selectedBytes")
  $summary.Add("$group.ActualPercentOfMax=$actualPercent")
}

$summary.Add('')
$summary.Add('SELECTED BY EXTENSION')
foreach($ext in ($script:SelectedByExtension.Keys | Sort-Object)){
  $summary.Add("$ext.Files=$($script:SelectedByExtension[$ext])")
  $summary.Add("$ext.Bytes=$($script:SelectedBytesByExtension[$ext])")
}

$summary.Add('')
$summary.Add("SamplePlan=$planPath")
$summary.Add("OversizeCandidates=$oversizePath")
$summary.Add("InterventionWindow=$interventionPath")
$summary.Add("AfterInterventionReview=$afterPath")
$summary.Add("OlderThanFiveYearsReview=$olderPath")
$summary.Add("ZeroByteCandidates=$zeroPath")
$summary.Add('')
$summary.Add('COPY_PERFORMED=False')

$summary | Set-Content -LiteralPath $summaryPath -Encoding UTF8
Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host
Write-Host 'FILE AUDIT SELECTOR V1.2 - DONE'
Write-Host '================================'
Write-Host ("Rows read              : {0:N0}" -f $RowsRead)
Write-Host ("Relevant candidates    : {0:N0}" -f $RelevantRows)
Write-Host ("Selected files         : {0:N0}" -f $script:SelectedFiles)
Write-Host ("Selected decimal GB    : {0:N3}" -f ($script:SelectedBytes/1000000000))
Write-Host ("Selected Desktop files : {0:N0}" -f $script:SelectedDesktopFiles)
Write-Host ("Source read-only       : {0}" -f $diskEnd.IsReadOnly)
Write-Host "Plan                    : $planPath"
Write-Host "Summary                 : $summaryPath"
Write-Host "Oversize                : $oversizePath"
Write-Host "Intervention            : $interventionPath"
Write-Host "Older than five years   : $olderPath"
Write-Host "Zero-byte               : $zeroPath"
Write-Host 'No source files were copied.'
