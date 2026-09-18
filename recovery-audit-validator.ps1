#requires -Version 5.1

<#
.SYNOPSIS
  Heuristically validate the approved sample contained in a mounted Recovery
  Audit Toolkit ISO.

.DESCRIPTION
  Reads approved-sample-plan.csv from a mounted read-only ISO and validates
  exactly those files. The ISO is never modified.

  Checks include:
    - presence and exact byte size;
    - header / magic bytes;
    - readable probes at BOF, 25%, 50%, 75%, and EOF;
    - format-aware checks for OOXML/ZIP, PDF, common images, WAV, PSD,
      Audacity AUP3/SQLite, JSON, legacy Office compound files, AI/PostScript,
      RAR, MP3/AAC, and text-like files.

  Results:
    PASS_STRONG  meaningful internal structure parsed successfully
    PASS_BASIC   size/signature/basic structure plausible
    WARN         suspicious but not conclusive
    FAIL         definite missing/truncated/unreadable/structurally invalid
    UNKNOWN      present/readable, but no reliable built-in parser implemented

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-validator.ps1" -IsoRoot "F:\"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidatePattern('^[A-Za-z]:\\?$')]
  [string]$IsoRoot,

  [Parameter()]
  [string]$ExpectedVolumeLabel='RECOVERY_AUDIT',

  [Parameter()]
  [string]$OutputDirectory='',

  [Parameter()]
  [switch]$CalculateSHA256
)

$ErrorActionPreference='Stop'
$ScriptVersion='1.0.0'
$Started=Get-Date
$RunID=$Started.ToString('yyyyMMdd_HHmmss')
$HeadBytes=65536
$TailBytes=65536
$ProbeBytes=16384

function Normalize-DriveRoot([string]$Path) {
  $p=[System.IO.Path]::GetFullPath($Path)
  if(-not $p.EndsWith('\')){$p+='\'}
  if($p -notmatch '^[A-Za-z]:\\$'){throw "IsoRoot must be a drive root such as F:\. Received: $Path"}
  return $p
}

function Convert-ToExtendedPath([string]$Path) {
  $full=[System.IO.Path]::GetFullPath($Path)
  if($full.StartsWith('\\?\')){return $full}
  if($full.StartsWith('\\')){return '\\?\UNC\'+$full.Substring(2)}
  return '\\?\'+$full
}

function Csv([AllowNull()][string]$Value) {
  if($null -eq $Value){return '""'}
  return '"' + $Value.Replace('"','""') + '"'
}

function Format-Duration([TimeSpan]$Span) {
  $h=[int][math]::Floor($Span.TotalHours)
  return ('{0:00}:{1:00}:{2:00}' -f $h,$Span.Minutes,$Span.Seconds)
}

function Open-ReadStream([string]$Path) {
  return New-Object System.IO.FileStream(
    (Convert-ToExtendedPath $Path),
    [System.IO.FileMode]::Open,
    [System.IO.FileAccess]::Read,
    [System.IO.FileShare]::ReadWrite,
    1048576,
    [System.IO.FileOptions]::SequentialScan
  )
}

function Test-FileExtended([string]$Path) {
  return [System.IO.File]::Exists((Convert-ToExtendedPath $Path))
}

function Get-FileLengthExtended([string]$Path) {
  $s=Open-ReadStream $Path
  try{return [UInt64]$s.Length}finally{$s.Dispose()}
}

function Read-Range([System.IO.FileStream]$Stream,[Int64]$Offset,[int]$Count) {
  if($Offset -lt 0){$Offset=0}
  if($Offset -gt $Stream.Length){$Offset=$Stream.Length}
  [void]$Stream.Seek($Offset,[System.IO.SeekOrigin]::Begin)
  $remaining=[int][math]::Min([Int64]$Count,($Stream.Length-$Offset))
  if($remaining -le 0){return [byte[]]@()}
  $buffer=New-Object byte[] $remaining
  $total=0
  while($total -lt $remaining){
    $n=$Stream.Read($buffer,$total,$remaining-$total)
    if($n -le 0){break}
    $total+=$n
  }
  if($total -eq $buffer.Length){return $buffer}
  $trimmed=New-Object byte[] $total
  if($total -gt 0){[Array]::Copy($buffer,$trimmed,$total)}
  return $trimmed
}

function Starts-WithBytes([byte[]]$Bytes,[byte[]]$Prefix) {
  if($null -eq $Bytes -or $Bytes.Length -lt $Prefix.Length){return $false}
  for($i=0;$i -lt $Prefix.Length;$i++){if($Bytes[$i] -ne $Prefix[$i]){return $false}}
  return $true
}

function Get-Ascii([byte[]]$Bytes) {
  if($null -eq $Bytes){return ''}
  return [System.Text.Encoding]::ASCII.GetString($Bytes)
}

function Get-U16BE([byte[]]$B,[int]$O) {
  return [UInt16](($B[$O] -shl 8) -bor $B[$O+1])
}

function Get-U16LE([byte[]]$B,[int]$O) {
  return [UInt16]($B[$O] -bor ($B[$O+1] -shl 8))
}

function Get-U32BE([byte[]]$B,[int]$O) {
  return [UInt32](([UInt32]$B[$O] -shl 24) -bor ([UInt32]$B[$O+1] -shl 16) -bor ([UInt32]$B[$O+2] -shl 8) -bor [UInt32]$B[$O+3])
}

function Get-U32LE([byte[]]$B,[int]$O) {
  return [UInt32](([UInt32]$B[$O]) -bor ([UInt32]$B[$O+1] -shl 8) -bor ([UInt32]$B[$O+2] -shl 16) -bor ([UInt32]$B[$O+3] -shl 24))
}

function Get-U64BE([byte[]]$B,[int]$O) {
  [UInt64]$v=0
  for($i=0;$i -lt 8;$i++){$v=($v -shl 8) -bor [UInt64]$B[$O+$i]}
  return $v
}

function New-V([string]$Result,[string]$Level,[string]$Detail) {
  return [pscustomobject]@{Result=$Result;Level=$Level;Detail=$Detail}
}

function Get-KeyValueFile([string]$Path) {
  $map=@{}
  $r=New-Object System.IO.StreamReader((Convert-ToExtendedPath $Path),[System.Text.Encoding]::UTF8,$true,65536)
  try{
    while(($line=$r.ReadLine()) -ne $null){
      if($line -match '^(?<Key>[^=]+)=(?<Value>.*)$'){$map[$matches.Key]=$matches.Value}
    }
  }finally{$r.Dispose()}
  return $map
}

function Find-SuspiciousRun([byte[]]$Bytes) {
  if($null -eq $Bytes -or $Bytes.Length -lt 8192){return ''}
  $value=$Bytes[0]
  $run=1
  for($i=1;$i -lt $Bytes.Length;$i++){
    if($Bytes[$i] -eq $value){
      $run++
      if($run -ge 8192 -and ($value -eq 0 -or $value -eq 255)){
        if($value -eq 0){return 'LONG_ZERO_RUN'}
        return 'LONG_FF_RUN'
      }
    }else{
      $value=$Bytes[$i]
      $run=1
    }
  }
  return ''
}

function Get-Probes([string]$Path,[UInt64]$Size) {
  $s=Open-ReadStream $Path
  try{
    $head=Read-Range $s 0 $HeadBytes
    $tail=Read-Range $s ([Int64][math]::Max(0,[Int64]$Size-$TailBytes)) $TailBytes
    $notes=New-Object System.Collections.Generic.List[string]
    foreach($p in @(
      @('BOF',[Int64]0),
      @('Q25',[Int64]([math]::Floor([double]$Size*0.25))),
      @('MID',[Int64]([math]::Floor([double]$Size*0.50))),
      @('Q75',[Int64]([math]::Floor([double]$Size*0.75))),
      @('EOF',[Int64][math]::Max(0,[Int64]$Size-$ProbeBytes))
    )){
      $b=Read-Range $s ([Int64]$p[1]) $ProbeBytes
      $flag=Find-SuspiciousRun $b
      if($flag){$notes.Add("$($p[0]):$flag")}
    }
    return [pscustomobject]@{Head=$head;Tail=$tail;Notes=($notes -join '|')}
  }finally{$s.Dispose()}
}

function Get-Magic([string]$Ext,[byte[]]$Head) {
  $e=$Ext.ToLowerInvariant()
  switch($e){
    '.docx' {if(Starts-WithBytes $Head ([byte[]](0x50,0x4B))){return 'MATCH'};return 'MISMATCH'}
    '.xlsx' {if(Starts-WithBytes $Head ([byte[]](0x50,0x4B))){return 'MATCH'};return 'MISMATCH'}
    '.pptx' {if(Starts-WithBytes $Head ([byte[]](0x50,0x4B))){return 'MATCH'};return 'MISMATCH'}
    '.zip'  {if(Starts-WithBytes $Head ([byte[]](0x50,0x4B))){return 'MATCH'};return 'MISMATCH'}
    '.pdf'  {if($Head.Length -ge 5 -and (Get-Ascii $Head[0..4]) -eq '%PDF-'){return 'MATCH'};return 'MISMATCH'}
    '.psd'  {if($Head.Length -ge 4 -and (Get-Ascii $Head[0..3]) -eq '8BPS'){return 'MATCH'};return 'MISMATCH'}
    '.jpg'  {if($Head.Length -ge 2 -and $Head[0] -eq 255 -and $Head[1] -eq 216){return 'MATCH'};return 'MISMATCH'}
    '.jpeg' {if($Head.Length -ge 2 -and $Head[0] -eq 255 -and $Head[1] -eq 216){return 'MATCH'};return 'MISMATCH'}
    '.png'  {if(Starts-WithBytes $Head ([byte[]](0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A))){return 'MATCH'};return 'MISMATCH'}
    '.tif'  {if((Starts-WithBytes $Head ([byte[]](0x49,0x49,0x2A,0x00))) -or (Starts-WithBytes $Head ([byte[]](0x4D,0x4D,0x00,0x2A)))){return 'MATCH'};return 'MISMATCH'}
    '.tiff' {if((Starts-WithBytes $Head ([byte[]](0x49,0x49,0x2A,0x00))) -or (Starts-WithBytes $Head ([byte[]](0x4D,0x4D,0x00,0x2A)))){return 'MATCH'};return 'MISMATCH'}
    '.wav'  {if($Head.Length -ge 12 -and (Get-Ascii $Head[0..3]) -eq 'RIFF' -and (Get-Ascii $Head[8..11]) -eq 'WAVE'){return 'MATCH'};return 'MISMATCH'}
    '.aup3' {if($Head.Length -ge 16 -and (Get-Ascii $Head[0..14]) -eq 'SQLite format 3' -and $Head[15] -eq 0){return 'MATCH'};return 'MISMATCH'}
    '.doc'  {if(Starts-WithBytes $Head ([byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1))){return 'MATCH'};return 'MISMATCH'}
    '.xls'  {if(Starts-WithBytes $Head ([byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1))){return 'MATCH'};return 'MISMATCH'}
    '.rar'  {if((Starts-WithBytes $Head ([byte[]](0x52,0x61,0x72,0x21,0x1A,0x07,0x00))) -or (Starts-WithBytes $Head ([byte[]](0x52,0x61,0x72,0x21,0x1A,0x07,0x01,0x00)))){return 'MATCH'};return 'MISMATCH'}
    '.rtf'  {if((Get-Ascii $Head).StartsWith('{\rtf')){return 'MATCH'};return 'MISMATCH'}
    '.mp3'  {
      if($Head.Length -ge 3 -and (Get-Ascii $Head[0..2]) -eq 'ID3'){return 'MATCH'}
      if($Head.Length -ge 2 -and $Head[0] -eq 255 -and (($Head[1] -band 224) -eq 224)){return 'MATCH'}
      return 'MISMATCH'
    }
    '.m4a'  {if($Head.Length -ge 12 -and (Get-Ascii $Head[4..7]) -eq 'ftyp'){return 'MATCH'};return 'MISMATCH'}
    '.aac'  {if($Head.Length -ge 2 -and $Head[0] -eq 255 -and (($Head[1] -band 246) -eq 240)){return 'MATCH'};return 'MISMATCH'}
    '.ai'   {
      $a=Get-Ascii $Head
      if($a.StartsWith('%PDF-') -or $a.StartsWith('%!PS-Adobe')){return 'MATCH'}
      return 'UNKNOWN'
    }
    '.txt' {return 'NOT_APPLICABLE'}
    '.md' {return 'NOT_APPLICABLE'}
    '.py' {return 'NOT_APPLICABLE'}
    '.csv' {return 'NOT_APPLICABLE'}
    '.json' {return 'NOT_APPLICABLE'}
    '.indd' {return 'NOT_IMPLEMENTED'}
    default {return 'UNKNOWN'}
  }
}

function Validate-ZipLike([string]$Path,[string]$Ext) {
  try{Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop}catch{}
  try{Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop}catch{}
  $fs=Open-ReadStream $Path
  $zip=$null
  try{
    $zip=[System.IO.Compression.ZipArchive]::new($fs,[System.IO.Compression.ZipArchiveMode]::Read,$false)
    $names=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $buffer=New-Object byte[] 65536
    $entries=0
    [UInt64]$expanded=0
    foreach($entry in $zip.Entries){
      $entries++
      [void]$names.Add($entry.FullName)
      if($entry.FullName.EndsWith('/')){continue}
      $es=$entry.Open()
      try{while(($n=$es.Read($buffer,0,$buffer.Length)) -gt 0){$expanded+=[UInt64]$n}}finally{$es.Dispose()}
    }
    if($Ext -eq '.zip'){return New-V 'PASS_STRONG' 'CONTAINER' "ZIP entries=$entries expandedBytes=$expanded"}

    $required=@('[Content_Types].xml')
    if($Ext -eq '.docx'){$required+=@('word/document.xml')}
    if($Ext -eq '.xlsx'){$required+=@('xl/workbook.xml')}
    if($Ext -eq '.pptx'){$required+=@('ppt/presentation.xml')}
    foreach($req in $required){if(-not $names.Contains($req)){return New-V 'FAIL' 'OOXML' "Required part missing: $req"}}

    $settings=New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing=[System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver=$null
    $xmlCount=0
    foreach($entry in $zip.Entries){
      if(-not $entry.FullName.EndsWith('.xml',[System.StringComparison]::OrdinalIgnoreCase)){continue}
      $xmlCount++
      $es=$entry.Open()
      try{
        $xr=[System.Xml.XmlReader]::Create($es,$settings)
        try{while($xr.Read()){}}
        finally{$xr.Dispose()}
      }finally{$es.Dispose()}
    }
    return New-V 'PASS_STRONG' 'OOXML' "entries=$entries xmlParsed=$xmlCount expandedBytes=$expanded"
  }catch{
    return New-V 'FAIL' 'CONTAINER' $_.Exception.Message
  }finally{
    if($null -ne $zip){$zip.Dispose()}
    $fs.Dispose()
  }
}

function Validate-Pdf([string]$Path,[UInt64]$Size,[byte[]]$Head,[byte[]]$Tail) {
  if(-not (Get-Ascii $Head).StartsWith('%PDF-')){return New-V 'FAIL' 'PDF' 'Missing PDF header'}
  $t=Get-Ascii $Tail
  if(-not $t.Contains('%%EOF')){return New-V 'FAIL' 'PDF' 'Missing EOF marker near end'}
  $m=[regex]::Matches($t,'startxref\s+(?<Offset>\d+)\s+%%EOF',[System.Text.RegularExpressions.RegexOptions]::Singleline)
  if($m.Count -eq 0){return New-V 'WARN' 'PDF' 'EOF present but startxref not found'}
  [UInt64]$o=0
  if(-not [UInt64]::TryParse($m[$m.Count-1].Groups['Offset'].Value,[ref]$o)){return New-V 'WARN' 'PDF' 'Could not parse startxref'}
  if($o -ge $Size){return New-V 'FAIL' 'PDF' "startxref beyond EOF: $o"}
  $s=Open-ReadStream $Path
  try{
    $p=(Get-Ascii (Read-Range $s ([Int64]$o) 128)).TrimStart()
    if($p.StartsWith('xref') -or $p -match '^\d+\s+\d+\s+obj'){return New-V 'PASS_STRONG' 'PDF' "header+EOF+startxref plausible offset=$o"}
    return New-V 'WARN' 'PDF' 'startxref target not recognised'
  }finally{$s.Dispose()}
}

function Validate-Image([string]$Path) {
  try{Add-Type -AssemblyName System.Drawing -ErrorAction Stop}catch{}
  $s=Open-ReadStream $Path
  $img=$null
  try{
    $img=[System.Drawing.Image]::FromStream($s,$true,$true)
    $w=$img.Width
    $h=$img.Height
    if($w -le 0 -or $h -le 0){return New-V 'FAIL' 'IMAGE' 'Invalid decoded dimensions'}
    return New-V 'PASS_STRONG' 'IMAGE' ("decoded={0}x{1}" -f $w,$h)
  }catch{
    return New-V 'FAIL' 'IMAGE' $_.Exception.Message
  }finally{
    if($null -ne $img){$img.Dispose()}
    $s.Dispose()
  }
}

function Validate-Wav([string]$Path,[UInt64]$Size) {
  $s=Open-ReadStream $Path
  try{
    $h=Read-Range $s 0 12
    if($h.Length -lt 12 -or (Get-Ascii $h[0..3]) -ne 'RIFF' -or (Get-Ascii $h[8..11]) -ne 'WAVE'){return New-V 'FAIL' 'RIFF' 'Invalid RIFF/WAVE header'}
    [UInt64]$declared=[UInt64](Get-U32LE $h 4)+8
    if($declared -gt $Size){return New-V 'FAIL' 'RIFF' "Declared size exceeds file: $declared"}
    [Int64]$pos=12
    $chunks=0
    while($pos+8 -le [Int64]$Size){
      $ch=Read-Range $s $pos 8
      if($ch.Length -lt 8){break}
      [UInt64]$len=Get-U32LE $ch 4
      [Int64]$next=$pos+8+[Int64]$len
      if(($len % 2) -eq 1){$next++}
      if($next -gt [Int64]$Size){return New-V 'FAIL' 'RIFF' "Chunk at $pos exceeds EOF"}
      $chunks++
      $pos=$next
    }
    return New-V 'PASS_STRONG' 'RIFF' "chunks=$chunks declaredBytes=$declared"
  }finally{$s.Dispose()}
}

function Validate-Psd([string]$Path,[UInt64]$Size) {
  $s=Open-ReadStream $Path
  try{
    $h=Read-Range $s 0 26
    if($h.Length -lt 26 -or (Get-Ascii $h[0..3]) -ne '8BPS'){return New-V 'FAIL' 'PSD' 'Invalid PSD header'}
    $version=Get-U16BE $h 4
    if($version -notin @(1,2)){return New-V 'FAIL' 'PSD' "Unsupported version=$version"}
    $channels=Get-U16BE $h 12
    [UInt32]$rows=Get-U32BE $h 14
    [UInt32]$cols=Get-U32BE $h 18
    $depth=Get-U16BE $h 22
    if($channels -lt 1 -or $rows -eq 0 -or $cols -eq 0 -or $depth -notin @(1,8,16,32)){return New-V 'FAIL' 'PSD' 'Implausible dimensions/channels/depth'}

    [Int64]$pos=26
    foreach($section in @('ColorModeData','ImageResources')){
      $b=Read-Range $s $pos 4
      if($b.Length -lt 4){return New-V 'FAIL' 'PSD' "Missing $section length"}
      [UInt64]$len=Get-U32BE $b 0
      $pos+=4
      if(([UInt64]$pos+$len) -gt $Size){return New-V 'FAIL' 'PSD' "$section exceeds EOF"}
      $pos+=[Int64]$len
    }

    if($version -eq 2){
      $b=Read-Range $s $pos 8
      if($b.Length -lt 8){return New-V 'FAIL' 'PSD' 'Missing LayerAndMask length'}
      [UInt64]$len=Get-U64BE $b 0
      $pos+=8
    }else{
      $b=Read-Range $s $pos 4
      if($b.Length -lt 4){return New-V 'FAIL' 'PSD' 'Missing LayerAndMask length'}
      [UInt64]$len=Get-U32BE $b 0
      $pos+=4
    }

    if(([UInt64]$pos+$len) -gt $Size){return New-V 'FAIL' 'PSD' 'LayerAndMask exceeds EOF'}
    return New-V 'PASS_STRONG' 'PSD' ("version={0} {1}x{2} channels={3} depth={4}" -f $version,$cols,$rows,$channels,$depth)
  }finally{$s.Dispose()}
}

function Validate-Aup3([UInt64]$Size,[byte[]]$Head) {
  if($Head.Length -lt 100 -or (Get-Ascii $Head[0..14]) -ne 'SQLite format 3' -or $Head[15] -ne 0){return New-V 'FAIL' 'SQLITE' 'SQLite header missing'}
  [UInt32]$pageSize=Get-U16BE $Head 16
  if($pageSize -eq 1){$pageSize=65536}
  $valid=($pageSize -eq 65536) -or ($pageSize -ge 512 -and $pageSize -le 32768 -and (($pageSize -band ($pageSize-1)) -eq 0))
  if(-not $valid){return New-V 'FAIL' 'SQLITE' "Invalid page size=$pageSize"}
  if(($Size % [UInt64]$pageSize) -ne 0){return New-V 'WARN' 'SQLITE' "Size not multiple of page size=$pageSize"}
  return New-V 'PASS_BASIC' 'SQLITE' "header plausible pageSize=$pageSize"
}

function Validate-Compound([UInt64]$Size,[byte[]]$Head) {
  if($Head.Length -lt 512 -or -not (Starts-WithBytes $Head ([byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)))){return New-V 'FAIL' 'CFBF' 'Compound file header missing'}
  $shift=Get-U16LE $Head 30
  $sector=[math]::Pow(2,$shift)
  if($sector -notin @(512,4096)){return New-V 'FAIL' 'CFBF' "Invalid sector size=$sector"}
  return New-V 'PASS_BASIC' 'CFBF' "header plausible sectorSize=$sector"
}

function Validate-Json([string]$Path,[UInt64]$Size) {
  if($Size -gt 100000000){return New-V 'WARN' 'JSON' 'Full parse skipped above 100 MB'}
  try{Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop}catch{}
  $r=New-Object System.IO.StreamReader((Convert-ToExtendedPath $Path),[System.Text.Encoding]::UTF8,$true,1048576)
  try{$text=$r.ReadToEnd()}finally{$r.Dispose()}
  try{
    $js=New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $js.MaxJsonLength=[int]::MaxValue
    $js.RecursionLimit=1000
    [void]$js.DeserializeObject($text)
    return New-V 'PASS_STRONG' 'JSON' 'Full JSON parse succeeded'
  }catch{
    return New-V 'FAIL' 'JSON' $_.Exception.Message
  }
}

function Validate-Text([string]$Ext,[byte[]]$Head,[byte[]]$Tail) {
  if($Ext -eq '.rtf'){
    if(-not (Get-Ascii $Head).StartsWith('{\rtf')){return New-V 'FAIL' 'TEXT' 'RTF header missing'}
    return New-V 'PASS_BASIC' 'TEXT' 'RTF header plausible'
  }
  $utf16=$false
  if($Head.Length -ge 2){$utf16=(($Head[0] -eq 255 -and $Head[1] -eq 254) -or ($Head[0] -eq 254 -and $Head[1] -eq 255))}
  if(-not $utf16){
    $total=$Head.Length+$Tail.Length
    $nul=0
    foreach($b in $Head){if($b -eq 0){$nul++}}
    foreach($b in $Tail){if($b -eq 0){$nul++}}
    if($total -gt 0 -and ([double]$nul/[double]$total) -gt 0.10){return New-V 'WARN' 'TEXT' 'High NUL-byte ratio'}
  }
  return New-V 'PASS_BASIC' 'TEXT' 'Text probes readable'
}

function Validate-ByFormat([string]$Path,[string]$Ext,[UInt64]$Size,[byte[]]$Head,[byte[]]$Tail) {
  $e=$Ext.ToLowerInvariant()
  switch($e){
    '.docx' {return Validate-ZipLike $Path $e}
    '.xlsx' {return Validate-ZipLike $Path $e}
    '.pptx' {return Validate-ZipLike $Path $e}
    '.zip' {return Validate-ZipLike $Path $e}
    '.pdf' {return Validate-Pdf $Path $Size $Head $Tail}
    '.jpg' {return Validate-Image $Path}
    '.jpeg' {return Validate-Image $Path}
    '.png' {return Validate-Image $Path}
    '.tif' {return Validate-Image $Path}
    '.tiff' {return Validate-Image $Path}
    '.wav' {return Validate-Wav $Path $Size}
    '.psd' {return Validate-Psd $Path $Size}
    '.aup3' {return Validate-Aup3 $Size $Head}
    '.doc' {return Validate-Compound $Size $Head}
    '.xls' {return Validate-Compound $Size $Head}
    '.json' {return Validate-Json $Path $Size}
    '.txt' {return Validate-Text $e $Head $Tail}
    '.md' {return Validate-Text $e $Head $Tail}
    '.py' {return Validate-Text $e $Head $Tail}
    '.csv' {return Validate-Text $e $Head $Tail}
    '.rtf' {return Validate-Text $e $Head $Tail}
    '.ai' {
      $a=Get-Ascii $Head
      if($a.StartsWith('%PDF-')){return Validate-Pdf $Path $Size $Head $Tail}
      if($a.StartsWith('%!PS-Adobe')){
        if((Get-Ascii $Tail).Contains('%%EOF')){return New-V 'PASS_BASIC' 'POSTSCRIPT' 'Header and EOF marker present'}
        return New-V 'WARN' 'POSTSCRIPT' 'Header present, EOF marker not found'
      }
      return New-V 'UNKNOWN' 'AI' 'Variant not recognised as PDF-compatible or PostScript'
    }
    '.rar' {if((Get-Magic $e $Head) -eq 'MATCH'){return New-V 'PASS_BASIC' 'RAR' 'RAR signature present'};return New-V 'FAIL' 'RAR' 'RAR signature missing'}
    '.mp3' {if((Get-Magic $e $Head) -eq 'MATCH'){return New-V 'PASS_BASIC' 'MP3' 'ID3/MPEG signature present'};return New-V 'WARN' 'MP3' 'No ID3/MPEG signature at BOF'}
    '.aac' {if((Get-Magic $e $Head) -eq 'MATCH'){return New-V 'PASS_BASIC' 'AAC' 'ADTS sync present'};return New-V 'WARN' 'AAC' 'No ADTS sync at BOF'}
    '.m4a' {if((Get-Magic $e $Head) -eq 'MATCH'){return New-V 'PASS_BASIC' 'ISOBMFF' 'ftyp box present'};return New-V 'WARN' 'ISOBMFF' 'ftyp box not found at expected position'}
    '.indd' {return New-V 'UNKNOWN' 'INDD' 'No dependency-free InDesign structural parser implemented'}
    default {return New-V 'UNKNOWN' 'GENERIC' 'No format-specific validator implemented'}
  }
}

function Get-SHA256([string]$Path) {
  $sha=[System.Security.Cryptography.SHA256]::Create()
  $s=Open-ReadStream $Path
  try{
    $hash=$sha.ComputeHash($s)
    return ([BitConverter]::ToString($hash)).Replace('-','').ToLowerInvariant()
  }finally{
    $s.Dispose()
    $sha.Dispose()
  }
}

$IsoRoot=Normalize-DriveRoot $IsoRoot
$driveLetter=$IsoRoot.Substring(0,1)
$volume=Get-Volume -DriveLetter $driveLetter -ErrorAction Stop

if(([string]$volume.DriveType) -ne 'CD-ROM'){throw "SAFETY ABORT: $IsoRoot is not mounted as CD-ROM. DriveType=$($volume.DriveType)"}
if($ExpectedVolumeLabel -and ([string]$volume.FileSystemLabel) -ne $ExpectedVolumeLabel){
  throw "SAFETY ABORT: volume label '$($volume.FileSystemLabel)' does not match '$ExpectedVolumeLabel'."
}

$PlanPath=Join-Path $IsoRoot 'approved-sample-plan.csv'
if(-not(Test-FileExtended $PlanPath)){throw "Approved sample plan not found: $PlanPath"}

$reconciled=Join-Path $IsoRoot 'copy-summary-reconciled.txt'
$original=Join-Path $IsoRoot 'copy-summary.txt'
if(Test-FileExtended $reconciled){$CopySummaryPath=$reconciled}
elseif(Test-FileExtended $original){$CopySummaryPath=$original}
else{throw 'No copy summary found on ISO.'}

$copySummary=Get-KeyValueFile $CopySummaryPath
if(-not $copySummary.ContainsKey('Status') -or $copySummary['Status'] -ne 'COMPLETE'){throw 'PRECHECK ABORT: copy summary is not COMPLETE.'}
if(-not $copySummary.ContainsKey('FailedFiles') -or [UInt64]$copySummary['FailedFiles'] -ne 0){throw 'PRECHECK ABORT: copy summary does not report FailedFiles=0.'}

if(-not $OutputDirectory){
  $desktop=[Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
  $OutputDirectory=Join-Path $desktop ('recovery-audit-validation_'+$RunID)
}else{
  $OutputDirectory=[System.IO.Path]::GetFullPath($OutputDirectory)
}

if(([System.IO.Path]::GetPathRoot($OutputDirectory)).TrimEnd('\').Equals($IsoRoot.TrimEnd('\'),[System.StringComparison]::OrdinalIgnoreCase)){
  throw 'SAFETY ABORT: validation output must not be written to ISO.'
}

if(Test-Path -LiteralPath $OutputDirectory){
  if(@(Get-ChildItem -LiteralPath $OutputDirectory -Force | Select-Object -First 1).Count -gt 0){throw "OutputDirectory is not empty: $OutputDirectory"}
}else{
  New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$ResultsPath=Join-Path $OutputDirectory 'validation-results.csv'
$ReviewPath=Join-Path $OutputDirectory 'validation-review.csv'
$SummaryPath=Join-Path $OutputDirectory 'validation-summary.txt'

$plan=Import-Csv -LiteralPath $PlanPath -Delimiter ';' -Encoding UTF8
if($null -eq $plan -or $plan.Count -eq 0){throw 'Approved sample plan is empty.'}

$required=@('PlanNo','FileID','RelativePath','Extension','SizeBytes','TypeGroup','TimeBand')
foreach($name in $required){if(-not ($plan[0].PSObject.Properties.Name -contains $name)){throw "Required plan column missing: $name"}}

[UInt64]$PlanBytes=0
foreach($row in $plan){
  [UInt64]$n=0
  if(-not [UInt64]::TryParse($row.SizeBytes,[ref]$n)){throw "Invalid SizeBytes for FileID $($row.FileID)"}
  $PlanBytes+=$n
}

if($copySummary.ContainsKey('PlanFiles') -and [UInt64]$copySummary['PlanFiles'] -ne [UInt64]$plan.Count){throw 'PRECHECK ABORT: plan row count differs from copy summary.'}
if($copySummary.ContainsKey('PlanBytes') -and [UInt64]$copySummary['PlanBytes'] -ne $PlanBytes){throw 'PRECHECK ABORT: plan byte total differs from copy summary.'}

$utf8=New-Object System.Text.UTF8Encoding($true)
$writer=New-Object System.IO.StreamWriter($ResultsPath,$false,$utf8,1048576)
$review=New-Object System.IO.StreamWriter($ReviewPath,$false,$utf8,262144)
$header='PlanNo;FileID;RelativePath;Extension;TypeGroup;TimeBand;ExpectedSizeBytes;ActualSizeBytes;Exists;SizeStatus;MagicStatus;BOFStatus;MiddleStatus;EOFStatus;ProbeNotes;ValidationLevel;Result;Detail;SHA256;DurationMs'
$writer.WriteLine($header)
$review.WriteLine($header)

$counts=@{PASS_STRONG=0;PASS_BASIC=0;WARN=0;FAIL=0;UNKNOWN=0}
$byExt=@{}
[UInt64]$processedFiles=0
[UInt64]$processedBytes=0
[UInt64]$missingFiles=0
[UInt64]$sizeMismatchFiles=0
$validationStarted=Get-Date
$lastUi=[datetime]::MinValue

Write-Host
Write-Host 'RECOVERY AUDIT VALIDATOR'
Write-Host '========================'
Write-Host "Version      : $ScriptVersion"
Write-Host "ISO          : $IsoRoot"
Write-Host "Volume label : $($volume.FileSystemLabel)"
Write-Host ("Plan files   : {0:N0}" -f $plan.Count)
Write-Host ("Plan bytes   : {0:N0} ({1:N3} decimal GB)" -f $PlanBytes,([double]$PlanBytes/1000000000))
Write-Host "Output       : $OutputDirectory"
Write-Host "SHA-256      : $([bool]$CalculateSHA256)"
Write-Host

try{
  foreach($row in $plan){
    $fileStart=Get-Date
    [UInt64]$expected=0
    [void][UInt64]::TryParse($row.SizeBytes,[ref]$expected)
    $rel=$row.RelativePath.TrimStart([char[]]@('\','/'))
    $path=[System.IO.Path]::GetFullPath((Join-Path $IsoRoot $rel))
    if(-not $path.StartsWith($IsoRoot,[System.StringComparison]::OrdinalIgnoreCase)){throw "RelativePath escapes ISO root: $($row.RelativePath)"}

    $exists=Test-FileExtended $path
    [UInt64]$actual=0
    $sizeStatus='NOT_CHECKED'
    $magic='NOT_CHECKED'
    $bof='NOT_CHECKED'
    $middle='NOT_CHECKED'
    $eof='NOT_CHECKED'
    $probeNotes=''
    $level='PRESENCE'
    $result='FAIL'
    $detail=''
    $hash=''

    if(-not $exists){
      $missingFiles++
      $detail='File missing from mounted ISO'
    }else{
      try{
        $actual=Get-FileLengthExtended $path
        if($actual -ne $expected){
          $sizeMismatchFiles++
          $sizeStatus='MISMATCH'
          $detail="Expected $expected bytes; found $actual"
        }else{
          $sizeStatus='MATCH'
          $p=Get-Probes $path $actual
          $bof='READABLE'
          $middle='READABLE'
          $eof='READABLE'
          $probeNotes=$p.Notes
          $magic=Get-Magic $row.Extension $p.Head
          $v=Validate-ByFormat $path $row.Extension $actual $p.Head $p.Tail
          $result=$v.Result
          $level=$v.Level
          $detail=$v.Detail

          if($magic -eq 'MISMATCH' -and $result -notin @('FAIL','UNKNOWN')){
            $result='WARN'
            $detail="Extension/header mismatch; $detail"
          }
          if($probeNotes -and $result -in @('PASS_BASIC','UNKNOWN')){
            $result='WARN'
            $detail="Suspicious sampled byte run: $probeNotes; $detail"
          }
          if($CalculateSHA256){$hash=Get-SHA256 $path}
        }
      }catch{
        $result='FAIL'
        $level='READ'
        $detail=$_.Exception.Message
      }
    }

    if(-not $counts.ContainsKey($result)){$result='UNKNOWN'}
    $counts[$result]++

    $ext=$row.Extension.ToLowerInvariant()
    if(-not $byExt.ContainsKey($ext)){$byExt[$ext]=@{Files=0;PASS_STRONG=0;PASS_BASIC=0;WARN=0;FAIL=0;UNKNOWN=0}}
    $byExt[$ext].Files++
    $byExt[$ext][$result]++

    $ms=[int][math]::Round(((Get-Date)-$fileStart).TotalMilliseconds)
    $line=@(
      (Csv $row.PlanNo),(Csv $row.FileID),(Csv $row.RelativePath),(Csv $row.Extension),(Csv $row.TypeGroup),(Csv $row.TimeBand),
      ([string]$expected),([string]$actual),(Csv ([string]$exists)),(Csv $sizeStatus),(Csv $magic),(Csv $bof),(Csv $middle),(Csv $eof),
      (Csv $probeNotes),(Csv $level),(Csv $result),(Csv $detail),(Csv $hash),([string]$ms)
    ) -join ';'
    $writer.WriteLine($line)
    if($result -in @('WARN','FAIL','UNKNOWN')){$review.WriteLine($line)}

    $processedFiles++
    $processedBytes+=$expected
    if(($processedFiles % 100) -eq 0){$writer.Flush();$review.Flush()}

    $now=Get-Date
    if(($now-$lastUi).TotalMilliseconds -ge 250 -or $processedFiles -eq $plan.Count){
      $elapsed=$now-$validationStarted
      $percent=if($PlanBytes -gt 0){([double]$processedBytes/[double]$PlanBytes)*100}else{100}
      $rate=if($elapsed.TotalSeconds -gt 0){[double]$processedBytes/$elapsed.TotalSeconds}else{0}
      $eta=[TimeSpan]::Zero
      if($rate -gt 0 -and $processedBytes -lt $PlanBytes){$eta=[TimeSpan]::FromSeconds(([double]($PlanBytes-$processedBytes))/$rate)}
      $status=('{0:N0}/{1:N0} files | {2:N3}/{3:N3} GB | {4:N1}% | ETA {5} | strong {6:N0} basic {7:N0} warn {8:N0} fail {9:N0} unknown {10:N0}' -f $processedFiles,$plan.Count,([double]$processedBytes/1000000000),([double]$PlanBytes/1000000000),$percent,(Format-Duration $eta),$counts.PASS_STRONG,$counts.PASS_BASIC,$counts.WARN,$counts.FAIL,$counts.UNKNOWN)
      $show=$path
      if($show.Length -gt 160){$show='...'+$show.Substring($show.Length-157)}
      Write-Progress -Activity 'Recovery Audit Toolkit - ISO file validation' -Status $status -CurrentOperation $show -PercentComplete ([math]::Max(0,[math]::Min(100,$percent)))
      $lastUi=$now
    }
  }
}finally{
  $writer.Flush();$writer.Dispose()
  $review.Flush();$review.Dispose()
  Write-Progress -Activity 'Recovery Audit Toolkit - ISO file validation' -Completed
}

$finished=Get-Date
$status=if($counts.FAIL -gt 0){'COMPLETE_WITH_FAILURES'}elseif($counts.WARN -gt 0 -or $counts.UNKNOWN -gt 0){'COMPLETE_WITH_FINDINGS'}else{'COMPLETE'}

$summary=New-Object System.Collections.Generic.List[string]
$summary.Add('RECOVERY AUDIT VALIDATOR')
$summary.Add('========================')
$summary.Add("Version=$ScriptVersion")
$summary.Add("Status=$status")
$summary.Add("Started=$($Started.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("Finished=$($finished.ToString('yyyy-MM-dd HH:mm:ss'))")
$summary.Add("Duration=$(Format-Duration ($finished-$Started))")
$summary.Add("IsoRoot=$IsoRoot")
$summary.Add("VolumeLabel=$($volume.FileSystemLabel)")
$summary.Add("DriveType=$($volume.DriveType)")
$summary.Add("ApprovedPlan=$PlanPath")
$summary.Add("CopySummary=$CopySummaryPath")
$summary.Add("PlanFiles=$($plan.Count)")
$summary.Add("PlanBytes=$PlanBytes")
$summary.Add("MissingFiles=$missingFiles")
$summary.Add("SizeMismatchFiles=$sizeMismatchFiles")
$summary.Add("PASS_STRONG=$($counts.PASS_STRONG)")
$summary.Add("PASS_BASIC=$($counts.PASS_BASIC)")
$summary.Add("WARN=$($counts.WARN)")
$summary.Add("FAIL=$($counts.FAIL)")
$summary.Add("UNKNOWN=$($counts.UNKNOWN)")
$summary.Add("CalculateSHA256=$([bool]$CalculateSHA256)")
$summary.Add("Results=$ResultsPath")
$summary.Add("Review=$ReviewPath")
$summary.Add('')
$summary.Add('BY EXTENSION')

foreach($ext in ($byExt.Keys | Sort-Object)){
  $x=$byExt[$ext]
  $summary.Add("$ext.Files=$($x.Files)")
  $summary.Add("$ext.PASS_STRONG=$($x.PASS_STRONG)")
  $summary.Add("$ext.PASS_BASIC=$($x.PASS_BASIC)")
  $summary.Add("$ext.WARN=$($x.WARN)")
  $summary.Add("$ext.FAIL=$($x.FAIL)")
  $summary.Add("$ext.UNKNOWN=$($x.UNKNOWN)")
}

$summary | Set-Content -LiteralPath $SummaryPath -Encoding UTF8

Write-Host
Write-Host 'RECOVERY AUDIT VALIDATOR - DONE'
Write-Host '==============================='
Write-Host "Status                  : $status"
Write-Host ("Validated files         : {0:N0}" -f $plan.Count)
Write-Host ("PASS_STRONG             : {0:N0}" -f $counts.PASS_STRONG)
Write-Host ("PASS_BASIC              : {0:N0}" -f $counts.PASS_BASIC)
Write-Host ("WARN                    : {0:N0}" -f $counts.WARN)
Write-Host ("FAIL                    : {0:N0}" -f $counts.FAIL)
Write-Host ("UNKNOWN                 : {0:N0}" -f $counts.UNKNOWN)
Write-Host ("Missing                 : {0:N0}" -f $missingFiles)
Write-Host ("Size mismatch           : {0:N0}" -f $sizeMismatchFiles)
Write-Host "Results                 : $ResultsPath"
Write-Host "Review                  : $ReviewPath"
Write-Host "Summary                 : $SummaryPath"
