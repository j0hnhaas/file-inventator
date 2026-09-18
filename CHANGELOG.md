# Changelog

This file records the development history of Recovery Audit Toolkit. The active repository root contains two executable tools: the inventory stage and the sampler stage.

## Unreleased — Recovery Audit Toolkit

- Added `recovery-audit-copy.ps1` as an optional controlled execution step for approved sample plans, with serial/read-only checks, no-overwrite safeguards, size verification, live progress, audit manifest, summary, and a preserved copy of the approved plan.
- Added `recovery-audit-iso.ps1` as a separate optional command that packages a completed sample folder into a UDF ISO image using Windows IMAPI2FS, with completion checks, a DVD-sized default limit, and byte-based write progress.

- Fixed a sampler reporting bug where summary-loop variables could overwrite the final console totals after the plan had already been built. The sample plan itself was unaffected.

- Repositioned the project from a single file inventory script to a two-stage recovery-audit toolkit.
- Active tools renamed to:
  - `recovery-audit-inventory.ps1`
  - `recovery-audit-sampler.ps1`
- README rewritten around the two-stage workflow rather than historical script versions.
- Earlier experimental sampler scripts removed from the active repository root; their history remains in Git.

## Sampler 1.2

- Added 14 mutually exclusive time strata from one day to five years before a documented failure event.
- Added global file-type targets across the complete sample.
- Prioritised native creative/project files:
  - `.psd`
  - `.ai`
  - `.indd`
  - `.aup3`
- Added three-stage sampling:
  1. time-by-type matrix;
  2. global type top-up;
  3. final priority fill.
- Added per-extension reporting.
- Preserved source serial-number and read-only checks.
- Kept oversize files, zero-byte files, files older than five years, intervention-window files, and post-intervention files in separate review lists.
- Planning remains non-copying.

## Sampler 1.1

- Replaced broad historical buckets with stratified age bands.
- Added per-time-band type reservations.
- Prioritised native creative and Office authoring formats.
- Removed smallest-file-first selection bias.
- Added richer sample summary output.

## Sampler 1.0

- Added first size-limited audit sample planner.
- Added Desktop prioritisation, recent-file prioritisation, oversize review, intervention-window review, and a default 3.9-billion-byte sample budget.
- Added repeated read-only verification during planning.

## Inventory 1.0.2

- Added verbose Robocopy output so skipped file entries can be represented.
- Tightened Robocopy error detection to avoid false positives from ordinary filenames.
- Preserved source serial-number and read-only verification.
- Master inventory remains hash-free and path-oriented.

## Inventory 1.0.1

- Fixed Robocopy argument binding.
- Added validated three-phase inventory workflow:
  1. pre-count;
  2. full Unicode listing;
  3. CSV parsing and validation.

## Inventory 1.0

- Initial read-only inventory workflow based on Robocopy list-only backup mode.
