<#
.SYNOPSIS
  Audits video files under a root folder and reports whether they match:
  - Video: HEVC Main10 + yuv420p10le or p010le
  - Audio: AAC (all audio streams)

.DESCRIPTION
  - Prompts for Root folder (unless -Root is provided).
  - Writes CSV report into a 'logs' folder next to this script (creates if missing).
  - Report columns:
    Path, VideoCodec, Profile, PixFmt, AudioCodecs, Action, Reason

.REQUIREMENTS
  - ffprobe must be available on PATH (FFmpeg install). :contentReference[oaicite:1]{index=1}

.EXAMPLE
  .\Audit-VideoStandard.ps1
  (prompts for root)

.EXAMPLE
  .\Audit-VideoStandard.ps1 -Root "D:\Media"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$Root,

  [Parameter(Mandatory = $false)]
  [string[]]$Extensions = @(".mkv", ".mp4", ".m4v", ".mov", ".avi")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-ScriptDir {
  if ($PSScriptRoot) { return $PSScriptRoot }
  # Fallback if invoked in unusual contexts
  return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

function Get-ToolPath {
  param([Parameter(Mandatory)][string]$Name)
  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required tool '$Name' not found on PATH. Install it and reopen PowerShell."
  }
  return $cmd.Path
}

function Get-FFprobeInfo {
  param(
    [Parameter(Mandatory)][string]$FFprobePath,
    [Parameter(Mandatory)][string]$FilePath
  )

  # Use JSON so we can reliably parse codec_name/profile/pix_fmt, etc. :contentReference[oaicite:2]{index=2}
  $ffprobeArgs = @(
    "-v","error",
    "-show_streams",
    "-show_format",
    "-of","json",
    $FilePath
  )

  $json = & $FFprobePath @ffprobeArgs 2>$null
  if (-not $json) { return $null }

  try {
    return $json | ConvertFrom-Json -ErrorAction Stop
  } catch {
    return $null
  }
}

function Get-ComplianceDecision {
  param(
    [Parameter(Mandatory)][object]$Probe
  )

  # Safely get streams array, handling null/empty cases
  if (-not $Probe -or -not $Probe.streams) {
    return @{
      VideoCodec  = ""
      Profile     = ""
      PixFmt      = ""
      AudioCodecs = ""
      Action      = "Error"
      Reason      = "No streams detected"
    }
  }

  $streams = @($Probe.streams)
  if ($streams.Count -eq 0) {
    return @{
      VideoCodec  = ""
      Profile     = ""
      PixFmt      = ""
      AudioCodecs = ""
      Action      = "Error"
      Reason      = "No streams detected"
    }
  }

  $v = $streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
  if (-not $v) {
    return @{
      VideoCodec  = ""
      Profile     = ""
      PixFmt      = ""
      AudioCodecs = ""
      Action      = "Error"
      Reason      = "No video stream detected"
    }
  }

  $audioStreams = @($streams | Where-Object { $_.codec_type -eq "audio" })
  $audioCodecs = if ($audioStreams -and $audioStreams.Count -gt 0) {
    ($audioStreams | ForEach-Object { $_.codec_name } | Where-Object { $_ } | Select-Object -Unique) -join ","
  } else {
    ""
  }

  $videoCodec = [string]$v.codec_name
  $videoProfile = [string]$v.profile
  $pixFmt     = [string]$v.pix_fmt

  $reasons = New-Object System.Collections.Generic.List[string]

  # Video checks (HEVC + Main 10 + yuv420p10le or p010le)
  # FFmpeg commonly reports this as "hevc", profile "Main 10", pix_fmt "yuv420p10le" or "p010le" (NVENC outputs p010le). :contentReference[oaicite:3]{index=3}
  if ($videoCodec -ne "hevc") { $reasons.Add("Video codec is '$videoCodec' (expected 'hevc')") }
  if ($videoProfile -ne "Main 10" -and $videoProfile -ne "Main10") { $reasons.Add("Profile is '$videoProfile' (expected 'Main 10')") }
  if ($pixFmt -ne "yuv420p10le" -and $pixFmt -ne "p010le") { $reasons.Add("PixFmt is '$pixFmt' (expected 'yuv420p10le' or 'p010le')") }

  # Audio checks (all audio streams must be AAC if any exist)
  if ($audioStreams -and $audioStreams.Count -gt 0) {
    $nonAac = @($audioStreams | Where-Object { $_.codec_name -ne "aac" })
    if ($nonAac -and $nonAac.Count -gt 0) {
      $bad = ($nonAac | ForEach-Object { $_.codec_name } | Select-Object -Unique) -join ","
      $reasons.Add("Audio codec(s) not AAC: $bad")
    }
  } else {
    # If you want "no audio" to be OK, remove this line.
    $reasons.Add("No audio streams detected")
  }

  if ($reasons.Count -eq 0) {
    return @{
      VideoCodec  = $videoCodec
      Profile     = $videoProfile
      PixFmt      = $pixFmt
      AudioCodecs = $audioCodecs
      Action      = "Keep"
      Reason      = "Already compliant"
    }
  }

  return @{
    VideoCodec  = $videoCodec
    Profile     = $videoProfile
    PixFmt      = $pixFmt
    AudioCodecs = $audioCodecs
    Action      = "Convert"
    Reason      = ($reasons -join " | ")
  }
}

# -------------------- Entry --------------------

$ffprobe = Get-ToolPath -Name "ffprobe"

if (-not $Root -or [string]::IsNullOrWhiteSpace($Root)) {
  $Root = Read-Host "Root folder to scan"
}

if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
  throw "Root folder not found: $Root"
}

$scriptDir = Resolve-ScriptDir
$logsDir   = Join-Path $scriptDir "logs"
if (-not (Test-Path -LiteralPath $logsDir -PathType Container)) {
  New-Item -ItemType Directory -Path $logsDir | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outCsv    = Join-Path $logsDir "VideoCodecAudit_$timestamp.csv"

$extSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($e in $Extensions) { [void]$extSet.Add($e) }

$files = Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction Stop |
  Where-Object { $extSet.Contains($_.Extension) }

$rows = foreach ($f in $files) {
  $probe = Get-FFprobeInfo -FFprobePath $ffprobe -FilePath $f.FullName
  if (-not $probe) {
    [pscustomobject]@{
      Path        = $f.FullName
      VideoCodec  = ""
      Profile     = ""
      PixFmt      = ""
      AudioCodecs = ""
      Action      = "Error"
      Reason      = "ffprobe failed or returned invalid JSON"
    }
    continue
  }

  $decision = Get-ComplianceDecision -Probe $probe

  [pscustomobject]@{
    Path        = $f.FullName
    VideoCodec  = $decision.VideoCodec
    Profile     = $decision.Profile
    PixFmt      = $decision.PixFmt
    AudioCodecs = $decision.AudioCodecs
    Action      = $decision.Action
    Reason      = $decision.Reason
  }
}

$rows | Export-Csv -LiteralPath $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "Report saved to: $outCsv"
Write-Host ("Files scanned: {0}" -f @($files).Count)
Write-Host ("Convert: {0} | Keep: {1} | Error: {2}" -f `
  (@($rows | Where-Object Action -eq "Convert").Count),
  (@($rows | Where-Object Action -eq "Keep").Count),
  (@($rows | Where-Object Action -eq "Error").Count)
)
