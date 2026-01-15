# MediaManager

Deterministic PowerShell scripts for auditing and standardizing a media library, optimized for **Cloudflare Tunnel delivery** and **Jellyfin compatibility**.

---

## Overview

MediaManager is a collection of **script-level, non-interactive PowerShell tools** designed to bring an existing media library into a **predictable, stream-safe state**.

While filenames, containers, and subtitles are standardized for **Jellyfin compatibility**, the **video and audio encoding standards are primarily driven by Cloudflare Tunnel constraints**, ensuring reliable remote streaming without transcoding surprises.

This repository represents the **stable foundation** of the MediaManager project.

---

## Design Principles

- **Audit-first**  
  Nothing is modified until the current state is explicitly measured.

- **Deterministic behavior**  
  No hidden state, no remembered paths, no background services.

- **Safety over speed**  
  All conversions use temp outputs and atomic swaps; originals are preserved.

- **Cloudflare-friendly encoding**  
  Encoding targets are chosen to minimize incompatibilities and unnecessary transcoding when media is accessed through Cloudflare Tunnels.

- **Analytics-ready outputs**  
  All reports and summaries are written as CSV files.

---

## Cloudflare Tunnel Considerations

When media is accessed through **Cloudflare Tunnels**, certain codecs and formats are significantly more reliable than others.

MediaManager standardizes on:
- **HEVC Main 10 video**
- **10-bit pixel formats**
- **AAC audio**
- **MKV container**

These choices:
- Reduce the likelihood of Cloudflare-side buffering or incompatibility
- Avoid edge-case transcoding paths
- Preserve quality while remaining bandwidth-efficient
- Remain fully compatible with Jellyfin clients

> Jellyfin compatibility is a requirement —  
> **Cloudflare Tunnel stability is the primary driver.**

---

## What This Repository Is

- A **script-only** toolset
- Safe to run on existing libraries
- Designed for **manual, intentional execution**
- Focused on **predictability and recoverability**

## What This Repository Is Not

- No GUI
- No background services or daemons
- No database
- No automatic file watchers
- No media scraping beyond explicit rename operations
- No subtitle analytics or reporting

---

## Included Scripts

| Script | Purpose |
|------|--------|
| `AuditVideoCodecs.ps1` | Audits media files and reports compliance via CSV |
| `Convert-And-Swap-To-HEVC10-AAC.ps1` | Safely converts non-compliant files and preserves originals |
| `MAL-Rename.ps1` | Renames anime episodes using MyAnimeList (Jikan API) |
| `SubtitleManager.ps1` | Integrates external subtitles into MKV files |

Each script is intentionally **single-purpose**.

---

## Governance & Design Contracts

This repository follows strict internal design contracts that define:
- media compliance rules
- safety guarantees (no data loss)
- stable CSV output schemas
- encoder selection and fallback behavior

These contracts are intentionally kept stable to ensure predictable behavior and safe iteration.

Changes that affect encoding standards, safety guarantees, or output schemas are considered breaking changes and are introduced deliberately.


---
