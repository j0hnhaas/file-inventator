# Recovery Audit Toolkit

**Recovery Audit Toolkit** is a two-stage PowerShell toolkit for evaluating a recovered, restored, or otherwise preserved Windows file system.

It does **not** perform data recovery itself. Its purpose is to make an existing recovery result inspectable and testable in a reproducible way.

The toolkit contains exactly two active scripts:

| Stage | Script | Purpose |
|---|---|---|
| 1 | `recovery-audit-inventory.ps1` | Build a read-only, ACL-aware master inventory of the mounted source volume |
| 2 | `recovery-audit-sampler.ps1` | Turn that inventory into a reproducible, size-limited, stratified audit sample plan |

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
- automatically prove that selected files can be opened;
- automatically copy the planned sample from the source.

Those functions are deliberately separate from the current read-only inventory and planning workflow.

## Reproducibility and limits

The toolkit is path-oriented. Hard-linked files may therefore appear once per enumerated path.

A single Excel worksheet is limited to 1,048,576 rows; large master CSVs should be analysed with PowerShell, Python, a database, or other tools that can handle larger datasets.

Sampling is a method for selecting files to inspect. It is not, by itself, proof of the quality or completeness of a recovery.

## Development status

The active codebase consists of the two scripts listed at the top of this README. Earlier experimental sampler versions remain available through Git history rather than as separate active tools in the repository root.

See [CHANGELOG.md](CHANGELOG.md) for development history.

## License

See [LICENSE](LICENSE).
