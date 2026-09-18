# Recovery Audit Toolkit

**Recovery Audit Toolkit** is a PowerShell toolkit for evaluating a recovered, restored, or otherwise preserved Windows file system. Its analytical core is deliberately split into two stages: inventory and stratified sample planning.

It does **not** perform data recovery itself. Its purpose is to make an existing recovery result inspectable and testable in a reproducible way.

The toolkit has **two core analysis scripts** plus three optional execution/validation utilities:

| Role | Script | Purpose |
|---|---|---|
| Core 1 | `recovery-audit-inventory.ps1` | Build a read-only, ACL-aware master inventory of the mounted source volume |
| Core 2 | `recovery-audit-sampler.ps1` | Turn that inventory into a reproducible, size-limited, stratified audit sample plan |
| Optional | `recovery-audit-copy.ps1` | Copy exactly the approved sample plan to a safe destination with progress and an audit manifest |
| Optional | `recovery-audit-iso.ps1` | Package a completed local sample folder into a UDF ISO image with its own explicit command |
| Optional | `recovery-audit-validator.ps1` | Heuristically validate exactly the approved sample files from a mounted ISO |

## Workflow

```text
Mounted Windows volume
        |
        | read-only
        v
+--------------------------------------+
| 1. INVENTORY                         |
| recovery-audit-inventory.ps1         |
|                                      |
| Enumerate the mounted file system    |
| with backup rights and create a      |
| validated master CSV.                |
+-------------------+------------------+
                    |
                    | master inventory
                    v
+--------------------------------------+
| 2. SAMPLE PLANNER                    |
| recovery-audit-sampler.ps1           |
|                                      |
| Stratify by age, file type, location |
| and byte budget to create an audit   |
| sample plan.                          |
+-------------------+------------------+
                    |
                    v
             sample-plan.csv
             + review lists
                    |
                    | optional
                    v
+--------------------------------------+
| 3. CONTROLLED COPY                   |
| recovery-audit-copy.ps1              |
|                                      |
| Copy exactly the approved plan to a  |
| safe local destination with progress |
| and an audit manifest.               |
+-------------------+------------------+
                    |
                    | optional, separate command
                    v
+--------------------------------------+
| 4. ISO PACKAGING                     |
| recovery-audit-iso.ps1               |
|                                      |
| Package the completed sample folder  |
| into a UDF ISO image.                |
+-------------------+------------------+
                    |
                    | mounted read-only ISO
                    v
+--------------------------------------+
| 5. ISO VALIDATION                    |
| recovery-audit-validator.ps1         |
|                                      |
| Verify approved files by size,       |
| BOF/mid/EOF probes, magic bytes and  |
| format-aware structural checks.      |
+--------------------------------------+
```

The separation is deliberate: the source-volume inventory can be expensive, but it only needs to be created once. Sampling strategies can then be rerun against the CSV without rescanning the source volume.

---

## Tool 1 — Inventory

### Purpose

`recovery-audit-inventory.ps1` answers:

> **Which files are present on this mounted volume, where are they, how large are they, and what timestamps do they carry?**

It creates the canonical master inventory used by Stage 2.

### How it works

The inventory tool uses **Robocopy in list-only backup mode**:

- `/L` — list only; no files are copied
- `/B` — enumerate with backup rights, including ACL-protected areas where permitted
- `/E` — include subdirectories
- `/XJ` — do not traverse junction targets
- `/SL` — list symbolic links rather than following their targets
- `/V` on the full listing — retain verbose per-file classifications, including skipped entries

The run has three phases:

1. **Pre-count** — determine expected file and byte totals.
2. **Raw listing** — create a Unicode Robocopy listing while showing progress.
3. **CSV + validation** — parse the listing into a structured master CSV and compare counts and bytes across the stages.

A fully validated run ends with:

```text
Status=COMPLETE
Validation=PASS
```

### Safety checks

The script refuses to start unless:

- the supplied source serial number matches the mounted disk;
- Windows reports the disk as `IsReadOnly=True`;
- the output Desktop is on a different volume.

The read-only state is checked again while the run is in progress.

> Software read-only is useful operational protection, but it is not a substitute for a hardware write blocker when strict forensic preservation is required.

### Main output

Each run creates a timestamped Desktop folder such as:

```text
recovery-audit-inventory_YYYYMMDD_HHMMSS/
├── recovery-audit-inventory_master.csv
├── recovery-audit-inventory_raw.txt
├── recovery-audit-inventory_precount.txt
├── recovery-audit-inventory_errors.txt
├── recovery-audit-inventory_parser-errors.txt
├── recovery-audit-inventory_progress.txt
└── recovery-audit-inventory_manifest.txt
```

The master CSV contains path, filename, extension, byte size, last-write timestamp, source identity, Robocopy classification, and related metadata. V1 does not calculate source-file content hashes.

---

## Tool 2 — Stratified sample planner

### Purpose

`recovery-audit-sampler.ps1` answers:

> **Which files should be inspected when the recovery contains far more files than can reasonably be opened manually?**

It reads the Stage 1 master CSV and produces a deterministic audit plan under a configurable byte limit.

The current sampler is **planning-only**. It does not yet copy selected source files and it does not claim that a file is healthy merely because it appears in the inventory.

### Time stratification

Files are placed into mutually exclusive age bands measured backwards from a documented failure event:

```text
1 day
3 days
1 week
2 weeks
3 weeks
1 month
6 weeks
2 months
3 months
6 months
1 year
2 years
3 years
5 years
```

Calendar months and years use calendar date arithmetic. Files older than five years are written to a separate review list rather than automatically consuming the main sample budget.

Files inside a defined intervention/backup window and files after that window are also reported separately.

### File-type prioritisation

The default global byte targets are:

```text
CREATIVE_NATIVE   35%   .psd .ai .indd .aup3
OFFICE_AUTHORING  20%   .docx .xlsx .xls .doc .pptx .rtf
MEDIA             20%   .jpg .jpeg .png .tif .tiff .wav .mp3 .m4a .aac
PDF               12%   .pdf
DATA_CODE          7%   .csv .json .py
TEXT_MARKUP        5%   .txt .md
ARCHIVE            1%   .zip .rar
```

Native creative/project files therefore receive the highest priority. Desktop files are preferred within comparable groups, and newer files are preferred over older ones.

### Sampling algorithm

The current sampler uses three planning phases:

1. **Time-by-type matrix** — reserve coverage across all time bands and file-type groups.
2. **Global type top-up** — move each type group toward its global byte target, with native creative files topped up first.
3. **Final priority fill** — use remaining budget in type-priority and recency order.

Oversize files, zero-byte files, files older than five years, intervention-window files, and post-intervention files are written to separate review CSVs.

The default maximum automatic sample is **3.9 billion bytes**, which can be changed with `-MaxSampleBytes`.

### Safety checks

The sampler verifies that:

- the supplied source serial number matches the mounted source disk;
- the source remains read-only;
- inventory/output are not located on the source volume;
- candidate rows belong to the expected source identity.

The read-only state is checked before and during both planning passes and again at completion.

### Main output

```text
recovery-audit-sampler_YYYYMMDD_HHMMSS/
├── sample-plan.csv
├── sample-summary.txt
├── oversize-candidates.csv
├── intervention-window.csv
├── after-intervention-review.csv
├── older-than-5-years-review.csv
└── zero-byte-candidates.csv
```

---

## Quick start

Run PowerShell **as Administrator**.

### 1. Create the inventory

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-inventory.ps1" `
  -SourceRoot "X:\" `
  -ExpectedSerial "<SERIAL_NUMBER>"
```

Check candidate disks first if needed:

```powershell
Get-Disk | Format-Table Number,FriendlyName,SerialNumber,Size,IsOffline,IsReadOnly
```

### 2. Build an audit sample plan

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-sampler.ps1" `
  -InventoryCsv "<PATH_TO_MASTER_CSV>" `
  -SourceRoot "X:\" `
  -ExpectedSerial "<SERIAL_NUMBER>" `
  -UserRoot "X:\Users\PROFILE\" `
  -FirstDocumentedFailureTime "YYYY-MM-DD HH:MM:SS" `
  -InterventionEndTime "YYYY-MM-DD HH:MM:SS" `
  -PlannedSampleFolderName "!AnalyseSampleRecovery"
```

Review `sample-summary.txt` and `sample-plan.csv` before any later copy or openability-testing step.

---

## Optional utility — Controlled copy

### Purpose

`recovery-audit-copy.ps1` executes an already approved `sample-plan.csv`. It does **not** select or reprioritise files.

It verifies the source-disk serial number and read-only state, rejects destination collisions and overwrites, preserves the planned relative directory structure, checks copied file lengths against the plan, and writes both `copy-manifest.csv` and `copy-summary.txt`.

The copy operation includes a live progress display with:

- processed files;
- processed gigabytes;
- percentage complete;
- throughput;
- estimated time remaining;
- failed-file count;
- current source path.

The approved sample plan itself is copied into the destination as `approved-sample-plan.csv` so the copied sample remains tied to the exact plan that created it.

Example:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-copy.ps1" `
  -SamplePlan "<PATH_TO_SAMPLE_PLAN>" `
  -SourceRoot "X:\" `
  -ExpectedSerial "<SERIAL_NUMBER>" `
  -DestinationRoot "C:\RecoveryAuditSample" `
  -ExpectedPlanFiles 18686 `
  -ExpectedPlanBytes 3900000000
```

---

## Optional utility — ISO packaging

### Purpose

`recovery-audit-iso.ps1` packages an already copied sample folder into an ISO file. ISO creation is **not automatic**; it is deliberately a separate, explicit command.

The utility uses the Windows Image Mastering API v2 and creates a single-session **UDF** image. It does not alter the source folder. If `copy-summary.txt` is present, it requires a completed copy with zero failed files unless the operator deliberately supplies `-AllowIncompleteCopy`.

The default maximum ISO size is 4,700,000,000 bytes. The image size is checked before the ISO is written. ISO writing has its own byte-based progress display with throughput and ETA.

Example:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-iso.ps1" `
  -SourceFolder "C:\RecoveryAuditSample" `
  -IsoPath "C:\RecoveryAuditSample.iso" `
  -VolumeLabel "RECOVERY_AUDIT"
```

---

## Optional utility — ISO validation

### Purpose

`recovery-audit-validator.ps1` validates the **mounted ISO itself**, not the original recovery volume or the intermediate copy folder.

It reads `approved-sample-plan.csv` from the ISO and therefore evaluates exactly the files selected by the sampler. Before validation it requires a CD-ROM mount, the expected volume label, a completed copy summary, zero failed copy files, and agreement between plan row/byte totals and the copy summary.

For every planned file it checks presence and exact byte size, then reads probes at the beginning of file, 25%, 50%, 75%, and end of file. It also checks file signatures and applies format-specific validation where feasible without external dependencies.

Current structural checks include OOXML/ZIP container traversal and XML parsing, PDF header/EOF/startxref checks, image decoding for JPEG/PNG/TIFF, RIFF/WAV chunk bounds, PSD/PSB section bounds, AUP3/SQLite header/page-size checks, JSON parsing, legacy Office compound-file headers, and basic checks for AI/PostScript, RAR, MP3, AAC/M4A and text-like files. InDesign files are currently reported as `UNKNOWN` after generic readability/size checks because no dependency-free structural parser is implemented.

Results are deliberately categorical rather than scored:

```text
PASS_STRONG  meaningful internal structure parsed successfully
PASS_BASIC   size/signature/basic structure plausible
WARN         suspicious but not conclusive
FAIL         definite missing/truncated/unreadable/structurally invalid
UNKNOWN      present/readable, but no reliable built-in parser implemented
```

The validator writes `validation-results.csv`, a focused `validation-review.csv` containing WARN/FAIL/UNKNOWN rows, and `validation-summary.txt` to a timestamped folder outside the ISO. It includes a live progress display with files, evaluated gigabytes, percentage, ETA, and running category counts.

Example for a mounted ISO at `F:\`:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-validator.ps1" -IsoRoot "F:\"
```

Per-file SHA-256 calculation is optional:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\recovery-audit-validator.ps1" -IsoRoot "F:\" -CalculateSHA256
```

A heuristic or structural PASS is evidence that the sampled file is present and structurally plausible. It is not a mathematical guarantee that every semantic element of the file is correct or that every application will open it successfully.

---

## What the toolkit does not do

Recovery Audit Toolkit currently does **not**:

- recover deleted files;
- scan unallocated space;
- carve file signatures;
- repair file systems;
- initialise or format disks;
- change source ownership or ACLs;
- inspect unmounted/RAW partitions;
- calculate source-file content hashes;
- prove semantic correctness or application-level openability with certainty;
- automatically copy the planned sample from the source without an explicit copy command;
- automatically create an ISO without an explicit ISO command.

Those functions are deliberately separate from the current read-only inventory and planning workflow.

## Reproducibility and limits

The toolkit is path-oriented. Hard-linked files may therefore appear once per enumerated path.

A single Excel worksheet is limited to 1,048,576 rows; large master CSVs should be analysed with PowerShell, Python, a database, or other tools that can handle larger datasets.

Sampling is a method for selecting files to inspect. It is not, by itself, proof of the quality or completeness of a recovery.

## Development status

The active codebase consists of two core analysis scripts and three optional execution/validation utilities listed at the top of this README. Earlier experimental sampler versions remain available through Git history rather than as separate active tools in the repository root.

See [CHANGELOG.md](CHANGELOG.md) for development history.

## License

See [LICENSE](LICENSE).
