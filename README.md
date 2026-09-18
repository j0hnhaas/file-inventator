# file-inventator

`file-inventator-v1.ps1` creates a read-only inventory of files on a mounted Windows volume.

The script is designed for situations where normal PowerShell enumeration may miss ACL-protected directories. It uses **Robocopy backup mode** for enumeration and keeps the source volume read-only.

## Goals

- fast, complete file-path inventory of the mounted volume
- no file-content hashing in V1
- no ownership or ACL changes
- no copying, deleting, formatting, initialization, or repair operations on the source
- Unicode raw listing for auditability
- semicolon-delimited UTF-8 CSV for analysis
- real percentage progress after an initial pre-count
- validation of file count and byte totals across the inventory stages

## Safety model

The script requires an elevated PowerShell session and refuses to start unless:

- the source is a drive root such as `X:\`
- the source disk serial number matches the serial supplied by the operator
- Windows reports the source disk as read-only
- the Desktop output location is on a different drive

Robocopy is invoked with `/L`, so it lists only. It does not copy or delete source files. Backup mode (`/B`) is used to enumerate ACL-protected areas without taking ownership or changing permissions.

> Software read-only is not a substitute for a hardware write blocker when strict forensic preservation is required.

## Scope

V1 inventories regular files reachable through the mounted source volume.

It does **not**:

- recover deleted files
- scan unallocated space
- carve file signatures
- inspect unmounted or RAW partitions
- calculate content hashes
- follow directory junction targets
- follow symbolic-link targets

Hard-linked files can appear once for each enumerated path because the inventory is path-oriented.

## Output

Each run creates a timestamped folder on the current user's Desktop:

```text
file-inventator-v1_YYYYMMDD_HHMMSS/
├── file-inventator-v1_master.csv
├── file-inventator-v1_raw.txt
├── file-inventator-v1_precount.txt
├── file-inventator-v1_errors.txt
├── file-inventator-v1_parser-errors.txt
├── file-inventator-v1_progress.txt
└── file-inventator-v1_manifest.txt
```

The master CSV contains:

```text
FileID
SourceID
FullPath
RelativePath
RelativeDirectory
FileName
BaseName
Extension
SizeBytes
LastWriteTime
LastWriteYear
TopLevelDirectory
PathLength
IsZeroByte
```

File IDs are simple sequential identifiers such as `F000000001`. V1 deliberately does not calculate SHA-256 hashes of source files.

## How it works

The run has three phases:

1. **Pre-count** – Robocopy enumerates the source in list-only backup mode and determines the expected total number of files and bytes.
2. **Raw listing** – a Unicode Robocopy listing is created. The script displays file-count progress, byte progress, throughput, ETA, and the current path.
3. **CSV + validation** – the raw listing is parsed into the master CSV. Counts and byte totals are compared between the pre-count, the full Robocopy run, and the CSV.

A successful run ends with:

```text
Status=COMPLETE
Validation=PASS
```

## Usage

Run PowerShell **as Administrator**.

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\file-inventator-v1.ps1" `
  -SourceRoot "X:\" `
  -ExpectedSerial "<SERIAL_NUMBER>"
```

To inspect the serial number before running the script:

```powershell
Get-Disk | Format-Table Number,FriendlyName,SerialNumber,Size,IsOffline,IsReadOnly
```

Confirm that the intended source disk reports `IsReadOnly = True` before starting.

## Excel note

A single Excel worksheet can display at most 1,048,576 rows. The CSV can contain more rows; use PowerShell, Python, a database, or split analysis files when the inventory is larger.

## Status

V1 is currently being tested against a large Windows file-system image. It should be treated as **pre-release until a full run completes with `Validation=PASS`**.
