# agents.md

## Project

This repo contains PowerShell scripts for auditing and standardizing a media library (Jellyfin-oriented).

Scripts (Phase 1, stable / source of truth):
- `Scripts/AuditVideoCodecs.ps1` — audit-only, read-only
- `Scripts/Convert-And-Swap-To-HEVC10-AAC.ps1` — conversion + backup + swap + summary CSV
- `Scripts/MAL-Rename.ps1` — rename episodes using MyAnimeList via Jikan; non-recursive; summary CSV
- `Scripts/SubtitleManager.ps1` — integrate external sidecar subtitles into MKVs using MKVToolNix (`mkvmerge`); non-recursive

Repo structure (current):
- `Docs/` — agent + planning docs (`agents.md`, `ToDo.md`)
- `Scripts/` — PowerShell scripts
- `Scripts/logs/` — **generated** logs (should be ignored by git)

---

## Standard

### Compliance definition (Audit + Convert target)
A video file is considered compliant when:
- Video codec: `hevc`
- Video profile: `Main 10` (or `Main10`)
- Video pix_fmt: `p010le` **or** `yuv420p10le`
  - `p010le` is accepted because NVENC commonly outputs it.
- Audio: **must have at least one audio stream**, and **all** audio streams must be `aac`
  - Current Phase-1 behavior: **no-audio is treated as non-compliant** (Audit marks `Convert`; Convert will attempt to standardize).
- Container after conversion: `.mkv`

---

## Goals
- Predictable scripts with minimal “unknown variables”.
- Audit produces a small, actionable CSV.
- Converter produces standardized `.mkv` files, preserves relative paths in backups, and avoids data loss.
- Summary CSVs are Power BI–compatible (`.csv`).

## Non-Goals
- No GUI work.
- No background services or daemons.
- No database.
- No subtitle inspection/reporting as analytics (SubtitleManager integrates only; it does not produce CSV reports).
- No “video quality tagging” (Jellyfin reads stream metadata).

---

# Script: `AuditVideoCodecs.ps1`

## Purpose
Scan a user-provided root folder and generate a CSV report indicating whether each file is compliant with the standard.

## Output
- Report is written to `./logs/` **relative to the script directory** (created if missing).
- CSV schema MUST remain exactly:

```
Path, VideoCodec, Profile, PixFmt, AudioCodecs, Action, Reason
```

## Semantics
- `Action`
  - `Keep` → already compliant
  - `Convert` → not compliant (includes “no audio streams detected”)
- `Reason`
  - For `Keep`: `Already compliant`
  - For `Convert`: one or more reasons joined with ` | `

## Required behavior (do not break)
- Audit must remain **read-only** (no file modifications).
- Do not reintroduce:
  - prompts for log folder paths other than the Root prompt
  - saved log directories / hidden state
- Do not add noisy debug logging unless explicitly requested.
- Do not add new external dependencies (e.g., `mkvmerge`).

## How to run
- Interactive (prompts for `Root`):
  - `./AuditVideoCodecs.ps1`
- Non-interactive:
  - `./AuditVideoCodecs.ps1 -Root "D:\Media"`

---

# Script: `Convert-And-Swap-To-HEVC10-AAC.ps1`

## Purpose
Convert non-compliant files into standardized `.mkv` (HEVC Main10 + AAC audio) and swap them in-place, while moving the original file to a backup location that preserves the relative folder structure.

## Required behavior (do not break)
- Output container MUST be `.mkv`.
- Must NOT skip files where:
  - the default audio stream is AAC **but**
  - another audio stream is non-AAC.
- Skip only when:
  - video is compliant **and**
  - **all** audio streams are AAC **and**
  - at least one audio stream exists (no-audio is treated as non-compliant).
- Conversion must be safe:
  - Encode to a temporary file first
  - Swap only after a successful encode
  - Move original to `BackupRoot` (preserving relative paths)

## Encoder policy
Preferred encoder order (per file):
1. `hevc_nvenc` (default)
2. `hevc_amf` (fallback)
3. `libx265` (CPU fallback)

Fallback is runtime-based:
- If NVENC fails for a file, try AMF
- If AMF fails, fall back to CPU x265

## Temp output naming
- Uses a temp output file alongside the source:
  - `*.mkv.__converting__.tmp.mkv`

## How to run
- Normal:
  - `./Convert-And-Swap-To-HEVC10-AAC.ps1 -MediaRoot "D:\Media" -BackupRoot "E:\Backups\Media"`
- Dry run:
  - `./Convert-And-Swap-To-HEVC10-AAC.ps1 -MediaRoot "D:\Media" -BackupRoot "E:\Backups\Media" -DryRun`

---

## Converter summary CSV (Power BI source)

### Purpose
Provide a Power BI–compatible summary of conversion outcomes.

### Location
- Written to `./logs/` relative to the script directory.
- Filename pattern:

```
ConvertSummary_YYYYMMDD_HHMMSS.csv
```

### Schema (MUST remain stable)
```
Status, Path, Title, Reason
```

### Semantics (matches current script behavior)
- `Status`
  - `Success` → file converted and swapped successfully
  - `Fail` → conversion failed after all encoder fallbacks
  - `DryRun` → logged for planned conversions during `-DryRun`
- `Path`
  - For `Success`: path to the final `.mkv`
  - For `Fail`: original file path
  - For `DryRun`: original file path
- `Title`
  - Filename **without** extension
- `Reason`
  - For `Success`: populated (e.g., `Converted successfully using hevc_nvenc`)
  - For `Fail`: populated (e.g., `All encoders failed (nvenc/amf/x265)...`)
  - For `DryRun`: populated (`Dry run - no conversion performed`)

### Rules
- CSV must be UTF-8 encoded.
- CSV is written once per run.

---

## Safety / data-loss rules
- Never delete originals as part of normal flow.
- Never overwrite final outputs except via the intended swap logic.
- On failure:
  - Clean up temp files
  - Log failure in summary CSV

---

# Script: `MAL-Rename.ps1`

## Purpose
Rename anime episode files to human-friendly titles using a MyAnimeList (MAL) series as the source of truth (episode list pulled via the Jikan API).

This script is intentionally **non-recursive**. It targets a single folder containing the episode files to rename.

## Inputs
- `-Season` (required): season number to embed in the output naming convention.
- `-Mal` (required): MAL numeric ID (e.g., `52588`) **or** MAL URL.
- `-Directory` (required): folder containing the files to rename (no subfolders).
- `-EpisodeOffset` (optional): subtracts from detected episode number (useful when files are offset vs MAL numbering).
- `-Ext` (optional): extensions to include.
- `-WhatIf` (optional): preview mode; will still produce the summary CSV (rows are written for planned renames).

## Episode matching rules (do not break)
- The script maps titles by parsing episode number from the filename, not by file ordering.
- Recognized patterns include (examples):
  - `S01E02`
  - `Ep 02` / `Episode 02`
  - `[02]`
- If episode number cannot be detected:
  - the file is skipped
  - a warning is printed
  - **no** summary row is written for that file (current behavior)
- If an episode number is detected but missing in Jikan data:
  - file is skipped with a warning (suggests adjusting `-EpisodeOffset`)
  - **no** summary row is written for that file (current behavior)

## Rename collision policy (matches current behavior)
- Safe and deterministic: never overwrite.
- If the target filename already exists:
  - the rename is skipped with a warning
  - **no** summary row is written for that file (current behavior)

## Jikan API behavior (matches current behavior)
- Uses Jikan episode endpoint with pagination.
- Includes a small fixed delay between requests (throttling).
- Does **not** currently implement exponential backoff retries; errors will surface rather than silently producing partial mappings.

## Output: summary CSV (Power BI friendly)
- Written to `./logs/` relative to the script directory.
- Filename pattern:

```
MALRenameSummary_YYYYMMDD_HHMMSS.csv
```

- Schema MUST remain stable:

```
Old Title, New Title
```

Notes (matches current behavior):
- Values are the full filenames **including extension** (e.g., `Episode 01.mkv`).
- Summary file is written only if at least one summary row exists.
  - If no renames happen (and not `-WhatIf`), the script prints “No renames performed; summary not written.”

## How to run
- Using MAL URL:
  - `./MAL-Rename.ps1 -Season 1 -Mal "https://myanimelist.net/anime/52588" -Directory "C:\Path\To\SeriesFolder"`
- Using MAL ID:
  - `./MAL-Rename.ps1 -Season 1 -Mal 52588 -Directory "C:\Path\To\SeriesFolder"`
- Preview:
  - `./MAL-Rename.ps1 -Season 1 -Mal 52588 -Directory "C:\Path\To\SeriesFolder" -WhatIf`

---

# Script: `SubtitleManager.ps1`

## Purpose
Integrate external subtitle sidecars (`.srt`, `.ass`, `.sup`, `.vtt`) into `.mkv` files using MKVToolNix `mkvmerge`.

This script is intentionally:
- Non-interactive (no prompts)
- Non-recursive (single folder only)
- “Safe swap” (temp output → verify → replace)

## Hard dependency
- `mkvmerge` must be installed and available in PATH (MKVToolNix).

## Inputs
- `-Directory` (required): folder containing `.mkv` files and sidecar subtitles.
- `-DefaultLanguage` (optional): language code used when no token is recognized in filename (default: `eng`).
- `-KeepBackup` (optional): keep a `.bak` copy of the original MKV after successful integration.
- `-NoDeleteSidecars` (optional): if set, sidecar files are kept (default behavior is to delete added sidecars).
- `-DryRun` (optional): prints planned actions without modifying files.

## Matching rules
- For each `.mkv`, sidecars are matched by:
  1) exact basename match or prefix match (`VideoName.*`)
  2) episode-number match (if an episode number is detected in the video filename)

## Language tagging rules (deterministic)
- Language is derived from filename tokens using boundary-aware regex patterns (e.g., `en/eng`, `ja/jpn`, etc.).
- Edge-cases like `ja[cc]` are treated as Japanese because brackets are considered valid token boundaries.

## Duplicate avoidance
- Existing subtitle tracks are detected via `mkvmerge --identify`.
- Sidecars whose language already exists in the MKV are skipped (by language code).

## Temp output naming / cleanup
- Temp output is created next to the source MKV:
  - `{BaseName}__converting__subs.tmp.mkv`
- Temp files are removed after successful swap, and also cleaned up on failure.

## Sidecar deletion (default)
- By default, sidecars that were successfully added are deleted.
- Use `-NoDeleteSidecars` to keep them.

## Safety / locked-file behavior
- If a video file is locked/in use (e.g., open in VLC), the script stops with a clear message and exit code `1`.

## Output
- Console output only (no CSV logs in Phase 1).

## How to run
- Normal:
  - `./SubtitleManager.ps1 -Directory "C:\Downloads\Series"`
- Dry run:
  - `./SubtitleManager.ps1 -Directory "C:\Downloads\Series" -DryRun`
- Default language override:
  - `./SubtitleManager.ps1 -Directory "C:\Downloads\Series" -DefaultLanguage jpn`

---

# Repo-wide change rules for agents
- Prefer minimal diffs and small, testable changes.
- Ask first if a change affects:
  - compliance rules
  - CSV schemas
  - encoder selection or fallback logic
  - backup/swap behavior
  - recursion/discovery behavior
- Do not reintroduce noisy debug/task logging unless explicitly requested.

## Coding conventions
- Keep functions small and single-purpose.
- Use `-LiteralPath` where applicable.
- Maintain:
  - `Set-StrictMode -Version Latest`
  - `$ErrorActionPreference = "Stop"` where used by the script
- Maintain UTF-8 compatibility for non-English paths and filenames.

## Validation checklist
Test with a small controlled set:
- Compliant: HEVC / Main 10 / p010le or yuv420p10le / ALL AAC
- Non-compliant video codec (e.g., h264)
- Non-compliant profile (Main, not Main 10)
- Non-compliant pix_fmt (yuv420p)
- Multiple audio streams where one is not AAC (must convert)
- No audio streams (should be flagged non-compliant per current behavior)
- Confirm backup preserves relative paths
- Confirm temp file usage and cleanup
- Confirm summary CSV is generated and Power BI–readable
- SubtitleManager:
  - sidecar matching by basename and by episode number
  - language detection boundaries (including bracket cases like `ja[cc]`)
  - duplicate language skip behavior
  - `-KeepBackup` and sidecar deletion behavior
