# agents.md

## Project
This repo contains PowerShell scripts for auditing and standardizing a media library for Jellyfin.

Scripts:
- `AuditVideoCodecs.ps1` (audit-only, read-only)
- `Convert-And-Swap-To-HEVC10-AAC.ps1` (conversion + backup + swap + summary logging)
- `MAL-Rename.ps1` (rename episodes using MyAnimeList via Jikan; non-recursive; summary logging)

These scripts are considered **source of truth**:
- Up to date
- Tested
- Safe for repeated use

`MAL-Rename.ps1` is also considered source of truth for episode-title renaming (see section below).

---

## Standard (Compliance Definition)
A file is considered compliant when:
- Video codec: HEVC (`hevc`)
- Video profile: Main 10 (`Main 10`)
- Video pix_fmt: **either** `p010le` **or** `yuv420p10le`
  - (`p010le` is accepted because NVENC commonly outputs it)
- Audio: **ALL** audio streams are AAC (`aac`)
- Container after conversion: `.mkv` (Jellyfin standard)

---

## Goals
- Predictable, low-noise scripts with minimal “unknown variables”.
- Audit produces a small, actionable CSV.
- Converter produces standardized `.mkv` files, preserves relative paths in backups, and avoids data loss.
- Logs are Power BI–compatible (`.csv`).

## Non-Goals
- No GUI work.
- No background services or daemons.
- No database.
- No subtitle inspection or reporting.
- No filename quality tagging is required for Jellyfin (single canonical encode workflow).

---

# Audit Script: `AuditVideoCodecs.ps1`

## Purpose
Scan a user-provided root folder and generate a CSV report indicating whether each file is compliant with the standard.

## Output
- Report is always written to `./logs/` relative to the script directory (create if missing).
- CSV schema MUST remain exactly:

```
Path, VideoCodec, Profile, PixFmt, AudioCodecs, Action, Reason
```

## Required Behavior (Do Not Break)
- Audit must remain **read-only** (no file modifications).
- Do not reintroduce:
  - prompts for log folder paths
  - saved log directories
- Do not add noisy debug or task logging unless explicitly requested.
- Do not add new external dependencies (e.g., `mkvmerge`) unless explicitly requested.

## How to Run
- Interactive:
  - Run script and enter root folder when prompted.
- Non-interactive:
  - `./AuditVideoCodecs.ps1 -Root "D:\Media"`

---

# Converter Script: `Convert-And-Swap-To-HEVC10-AAC.ps1`

## Purpose
Convert non-compliant files into a standardized `.mkv` format and swap them in-place, while moving the original file to a backup location that preserves the relative folder structure.

## Required Behavior (Do Not Break)
- Output container MUST be `.mkv`.
- Must NOT skip files where:
  - the default audio stream is AAC **but**
  - another audio stream is non-AAC.
- Skip only when **video is compliant AND all audio streams are AAC**.
- Conversion must be safe:
  - Encode to a temporary file first
  - Swap only after a successful encode
  - Always move original to `BackupRoot` (preserving relative paths)

## Encoder Policy
Preferred encoder order (per file):
1. `hevc_nvenc` (default)
2. `hevc_amf` (fallback)
3. `libx265` (CPU fallback)

- Fallback must be runtime-based:
  - If NVENC fails for a file, try AMF
  - If AMF fails, fall back to CPU x265

## How to Run
- Normal:
  - `./Convert-And-Swap-To-HEVC10-AAC.ps1 -MediaRoot "D:\Media" -BackupRoot "E:\Backups\Media"`
- Dry run:
  - `./Convert-And-Swap-To-HEVC10-AAC.ps1 -MediaRoot "D:\Media" -BackupRoot "E:\Backups\Media" -DryRun`

---

## Converter Summary CSV (Power BI Source)

### Purpose
Provide a Power BI–compatible summary of conversion outcomes.

### Location
- Always written to `./logs/` relative to the script directory.
- Filename pattern:

```
ConvertSummary_YYYYMMDD_HHMMSS.csv
```

### Schema (MUST remain stable)

```
Status, Path, Title, Reason
```

### Semantics
- `Status`
  - `Success` → file converted and swapped successfully
  - `Fail` → conversion failed after all encoder fallbacks
  - `DryRun` → logged only if DryRun logging is enabled
- `Path`
  - For `Success`: path to the final `.mkv`
  - For `Fail`: original file path
- `Title`
  - Filename without extension
- `Reason`
  - Empty for `Success`
  - Populated for `Fail` (e.g., “All encoders failed (nvenc/amf/x265)”)

### Rules
- CSV must be UTF-8 encoded.
- CSV is written once per run (not incremental per file).
- This CSV is the **primary structured log** for analytics (Power BI).

---

## Safety / Data Loss Rules
- Never delete originals as part of normal flow.
- Never overwrite final outputs except via the intended swap logic.
- On failure:
  - Clean up temp files
  - Log failure in summary CSV
  - (Optional) append to legacy fail list if still enabled

---


---

# Renamer Script: `MAL-Rename.ps1`

## Purpose
Rename anime episode files to human-friendly titles using a MyAnimeList (MAL) series as the source of truth (episode list pulled via the Jikan API).

This script is intentionally **non-recursive**. It targets a single folder containing the episode files to rename.

## Inputs
- `-Season` (required): season number to embed in the output naming convention.
- `-Mal` (required): MAL **numeric ID** (e.g., `52588`) **or** MAL URL (e.g., `https://myanimelist.net/anime/52588/...`).
- `-Directory` (required): folder containing the files to rename (no subfolders).
- `-EpisodeOffset` (optional): subtracts from detected episode number (useful when files are offset vs MAL numbering).
- `-Ext` (optional): extensions to include (default is video-centric; keep aligned with your library conventions).
- `-WhatIf` (optional): preview mode; MUST still produce the summary CSV.

## Episode Matching Rules (Do Not Break)
- The script MUST map titles by **parsing episode number from the filename**, not by file ordering.
- Priority match includes (examples):
  - `S01E02`
  - `Ep 02` / `Episode 02`
  - `[02]`
- If episode number cannot be detected for a file, it MUST:
  - skip the rename
  - write a summary row showing the old title and the unchanged new title (or leave new title blank), depending on current implementation
  - avoid guessing based on alphabetical order

## Rename Collision Policy (Do Not Break)
- Must be deterministic and safe.
- If the target filename already exists, the script MUST avoid overwriting.
- Preferred approach: append ` (2)`, ` (3)`, ... up to a reasonable limit and log the final chosen name.

## API Stability Rules (Jikan)
- Jikan calls MUST be resilient:
  - retry on transient failures (network, 429, 5xx)
  - exponential backoff (small, bounded)
- Failures must raise a clear error message (do not silently produce partial mappings).

## Output: Summary CSV (Power BI Friendly)
- Always written to `./logs/` relative to the script directory.
- Filename pattern:

```
MALRenameSummary_YYYYMMDD_HHMMSS.csv
```

- Schema MUST remain stable:

```
Old Title, New Title
```

Notes:
- `Old Title`: original filename without extension.
- `New Title`: final filename without extension (after collision handling).

## How to Run
- Using MAL URL:
  - `./MAL-Rename.ps1 -Season 1 -Mal "https://myanimelist.net/anime/52588" -Directory "C:\Path\To\SeriesFolder"`
- Using MAL ID:
  - `./MAL-Rename.ps1 -Season 1 -Mal 52588 -Directory "C:\Path\To\SeriesFolder"`
- Preview:
  - `./MAL-Rename.ps1 -Season 1 -Mal 52588 -Directory "C:\Path\To\SeriesFolder" -WhatIf`


# Repo-Wide Change Rules for Agents
- Prefer minimal diffs and small, testable changes.
- ASK FIRST if a change affects:
  - compliance rules
  - CSV schemas
  - encoder selection or fallback logic
  - backup/swap behavior
  - recursion or discovery behavior
- Do not reintroduce debug/task logging noise unless explicitly requested.

## Coding Conventions
- Keep functions small and single-purpose.
- Use `-LiteralPath` where applicable.
- Keep:
  - `Set-StrictMode -Version Latest`
  - `$ErrorActionPreference = "Stop"`
- Maintain UTF-8 compatibility for non-English paths and filenames.

## Validation Checklist
Test with a small controlled set:
- Compliant: HEVC / Main 10 / p010le or yuv420p10le / ALL AAC
- Non-compliant video codec (e.g., h264)
- Non-compliant profile (Main, not Main 10)
- Non-compliant pix_fmt (yuv420p)
- Multiple audio streams where one is not AAC (must convert)
- No audio streams (confirm expected behavior)
- Confirm backup preserves relative paths
- Confirm temp file usage and cleanup
- Confirm summary CSV is generated and Power BI–readable

