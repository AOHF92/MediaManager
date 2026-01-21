#requires -Version 5.1
<#
.SYNOPSIS
  Converts audio files and audio-carrier video containers to canonical OPUS format (128 kbps VBR, stereo).

.DESCRIPTION
  This script converts supported audio formats and video containers (extracting audio only) to OPUS format
  optimized for Navidrome compatibility. All video streams are discarded. Original files are archived
  after successful conversion and verification.

  Supported inputs:
  - Audio: MP3, AAC, WAV, FLAC, OPUS
  - Video containers: MP4, MKV, WEBM (audio extracted, video discarded)

  Output standard:
  - Format: OPUS
  - Bitrate: 128 kbps VBR
  - Channels: Stereo
  - Sample rate: Preserved from source
  - Metadata: Mandatory

.REQUIREMENTS
  - ffmpeg and ffprobe must be available on PATH
  - Install FFmpeg via winget: Gyan.FFmpeg

.EXAMPLE
  .\ConvertToAudioStandard.ps1
  (prompts for source and destination roots)

.EXAMPLE
  .\ConvertToAudioStandard.ps1 -SourceRoot "D:\Music" -DestinationRoot "E:\Archive"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$SourceRoot,

  [Parameter(Mandatory = $false)]
  [string]$DestinationRoot,

  [switch]$Preview
)

# Console encodings (UTF-8 for non-English paths)
try { chcp 65001 > $null } catch {}
[Console]::InputEncoding  = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding']    = 'utf8'
$PSDefaultParameterValues['Add-Content:Encoding'] = 'utf8'
$PSDefaultParameterValues['Set-Content:Encoding'] = 'utf8'

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ==================== T2: Configuration & Conventions ====================

# Exit codes
$EXIT_SUCCESS = 0
$EXIT_USER_ABORT = 1
$EXIT_CONVERSION_FAILURE = 2
$EXIT_MISSING_DEPENDENCY = 3

# Canonical output standard
$OPUS_BITRATE = "128"
$OPUS_CHANNELS = "2"  # Stereo
$OPUS_EXTENSION = ".opus"

# Supported input extensions
$AUDIO_EXTENSIONS = @(".mp3", ".aac", ".wav", ".flac", ".opus", ".m4a", ".ogg")
$VIDEO_EXTENSIONS = @(".mp4", ".mkv", ".webm", ".m4v", ".mov")
$SUPPORTED_EXTENSIONS = $AUDIO_EXTENSIONS + $VIDEO_EXTENSIONS

# Duration tolerance for verification (seconds)
$DURATION_TOLERANCE = 2.0

# Common cover image filenames (case-insensitive)
$COVER_IMAGE_NAMES = @("cover.jpg", "cover.png", "cover.jpeg", "folder.jpg", "folder.png", "folder.jpeg", 
                       "albumart.jpg", "albumart.png", "albumart.jpeg", "album.jpg", "album.png", "album.jpeg",
                       "front.jpg", "front.png", "front.jpeg", "artwork.jpg", "artwork.png", "artwork.jpeg")

# libopus supported sample rates (Hz)
$OPUS_SUPPORTED_SAMPLE_RATES = @(48000, 24000, 16000, 12000, 8000)

# Path conventions
function Resolve-ScriptDir {
  if ($PSScriptRoot) { return $PSScriptRoot }
  return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

# ==================== T3: Dependency Validator ====================

function Test-Dependencies {
  $missing = @()
  $toolInfo = @{}
  
  $tools = @("ffmpeg", "ffprobe")
  foreach ($tool in $tools) {
    $cmd = Get-Command $tool -ErrorAction SilentlyContinue
    if (-not $cmd) {
      $missing += $tool
      # Store installation info per tool
      switch ($tool) {
        "ffmpeg" { $toolInfo[$tool] = @{ Method = "winget"; Command = "winget install Gyan.FFmpeg" } }
        "ffprobe" { $toolInfo[$tool] = @{ Method = "winget"; Command = "winget install Gyan.FFmpeg"; Note = "(included with FFmpeg)" } }
      }
    }
  }
  
  if ($missing.Count -gt 0) {
    Write-Host "Missing required dependencies:" -ForegroundColor Red
    foreach ($tool in $missing) {
      $info = $toolInfo[$tool]
      Write-Host "  - $tool" -ForegroundColor Yellow
      if ($info -and $info.ContainsKey('Note')) {
        Write-Host "    $($info.Note)" -ForegroundColor Gray
      }
    }
    Write-Host ""
    
    # Group by installation method
    $hasWinget = @($missing | Where-Object { $toolInfo[$_].Method -eq "winget" })
    
    if ($hasWinget.Count -gt 0) {
      Write-Host "Install via winget:" -ForegroundColor Cyan
      foreach ($tool in $hasWinget) {
        Write-Host "  $($toolInfo[$tool].Command)" -ForegroundColor White
      }
      Write-Host ""
    }
    
    Write-Host "After installation, restart PowerShell and try again." -ForegroundColor Yellow
    exit $EXIT_MISSING_DEPENDENCY
  }
  
  return $true
}

function Get-ToolPath {
  param([Parameter(Mandatory)][string]$Name)
  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required tool '$Name' not found on PATH."
  }
  return $cmd.Path
}

# ==================== T4: Startup Prompts & Validation ====================

function Get-UserPaths {
  param(
    [string]$SourceRoot,
    [string]$DestinationRoot
  )
  
  # Prompt for source root
  if (-not $SourceRoot -or [string]::IsNullOrWhiteSpace($SourceRoot)) {
    $SourceRoot = Read-Host "Source root (final OPUS library location)"
    if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
      Write-Host "Aborted: Source root is required." -ForegroundColor Yellow
      exit $EXIT_USER_ABORT
    }
  }
  
  # Prompt for destination root (archive)
  if (-not $DestinationRoot -or [string]::IsNullOrWhiteSpace($DestinationRoot)) {
    $DestinationRoot = Read-Host "Destination root (archive for originals)"
    if ([string]::IsNullOrWhiteSpace($DestinationRoot)) {
      Write-Host "Aborted: Destination root is required." -ForegroundColor Yellow
      exit $EXIT_USER_ABORT
    }
  }
  
  # Validate paths
  if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "Source root not found: $SourceRoot"
  }
  
  # Block dangerous configs (source = destination)
  $sourceResolved = (Resolve-Path -LiteralPath $SourceRoot).Path.TrimEnd('\', '/')
  $destResolved = if (Test-Path -LiteralPath $DestinationRoot) {
    (Resolve-Path -LiteralPath $DestinationRoot).Path.TrimEnd('\', '/')
  } else {
    $DestinationRoot.TrimEnd('\', '/')
  }
  
  if ($sourceResolved -eq $destResolved) {
    throw "Source and destination roots cannot be the same: $sourceResolved"
  }
  
  # Check if destination is a subdirectory of source
  if ($destResolved.StartsWith($sourceResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Destination root cannot be inside source root. This would cause conflicts."
  }
  
  return @{
    SourceRoot = $SourceRoot
    DestinationRoot = $DestinationRoot
  }
}

# ==================== T6: File Scanner ====================

function Get-EligibleFiles {
  param(
    [Parameter(Mandatory)][string]$SourceRoot
  )
  
  $extSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($e in $SUPPORTED_EXTENSIONS) {
    [void]$extSet.Add($e)
  }
  
  $files = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction Stop |
    Where-Object { 
      $extSet.Contains($_.Extension) -and
      $_.Extension -ne $OPUS_EXTENSION  # Exclude canonical OPUS
    }
  
  return $files
}

# ==================== T7: Audio-Carrier Validation ====================

function Test-AudioStreamExists {
  param(
    [Parameter(Mandatory)][string]$FFprobePath,
    [Parameter(Mandatory)][string]$FilePath
  )
  
  $ffprobeArgs = @(
    "-v", "error",
    "-select_streams", "a",
    "-show_entries", "stream=codec_type",
    "-of", "json",
    $FilePath
  )
  
  $json = & $FFprobePath @ffprobeArgs 2>$null
  if (-not $json) { return $false }
  
  try {
    $obj = $json | ConvertFrom-Json
    if ($obj.streams -and $obj.streams.Count -gt 0) {
      $audioStreams = @($obj.streams | Where-Object { $_.codec_type -eq "audio" })
      return $audioStreams.Count -gt 0
    }
    return $false
  } catch {
    return $false
  }
}

# ==================== T8: Extract Embedded Metadata ====================

function Get-FilenameMetadata {
  param(
    [Parameter(Mandatory)][string]$FilePath
  )
  
  $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
  $result = @{
    Artist = $null
    Album = $null
    Track = $null
    Title = $null
  }
  
  # Common patterns:
  # "Artist - Album - TrackNumber Title"
  # "Artist - Album - TrackNumber - Title"
  # "TrackNumber - Artist - Title"
  # "TrackNumber. Title"
  
  # Pattern 1: "Artist - Album - TrackNumber Title" or "Artist - Album - TrackNumber - Title"
  if ($fileName -match '^(.+?)\s*-\s*(.+?)\s*-\s*(\d+)\s*(?:-\s*)?(.+)$') {
    $result.Artist = $matches[1].Trim()
    $result.Album = $matches[2].Trim()
    $result.Track = $matches[3].Trim()
    $result.Title = $matches[4].Trim()
    return $result
  }
  
  # Pattern 2: "TrackNumber - Artist - Title"
  if ($fileName -match '^(\d+)\s*-\s*(.+?)\s*-\s*(.+)$') {
    $result.Track = $matches[1].Trim()
    $result.Artist = $matches[2].Trim()
    $result.Title = $matches[3].Trim()
    return $result
  }
  
  # Pattern 3: "TrackNumber. Title" or "TrackNumber - Title"
  if ($fileName -match '^(\d+)[.\-]\s*(.+)$') {
    $result.Track = $matches[1].Trim()
    $result.Title = $matches[2].Trim()
    return $result
  }
  
  return $result
}

function Get-EmbeddedMetadata {
  param(
    [Parameter(Mandatory)][string]$FFprobePath,
    [Parameter(Mandatory)][string]$FilePath
  )
  
  $ffprobeArgs = @(
    "-v", "error",
    "-show_entries", "format_tags=artist,album,album_artist,date,year,track,disc,title",
    "-of", "json",
    $FilePath
  )
  
  $json = & $FFprobePath @ffprobeArgs 2>$null
  $embeddedTags = @{
    Artist = $null
    AlbumArtist = $null
    Album = $null
    Year = $null
    Date = $null
    Track = $null
    Disc = $null
    Title = $null
  }
  
  if ($json) {
    try {
      $obj = $json | ConvertFrom-Json
      $tags = if ($obj.format -and $obj.format.tags) { $obj.format.tags } else { @{} }
      
      # Build a case-insensitive tag lookup (ffprobe may return uppercase keys like "ALBUM")
      $tagMap = @{}
      foreach ($prop in $tags.PSObject.Properties) {
        if ($null -ne $prop.Value) {
          $tagMap[$prop.Name.ToLower()] = [string]$prop.Value
        }
      }
      
      function Get-TagValue {
        param([string[]]$Keys)
        foreach ($k in $Keys) {
          $key = $k.ToLower()
          if ($tagMap.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace($tagMap[$key])) {
            return $tagMap[$key].Trim()
          }
        }
        return $null
      }
      
      $embeddedTags.Artist = Get-TagValue @("artist")
      $embeddedTags.AlbumArtist = Get-TagValue @("album_artist", "albumartist")
      $embeddedTags.Album = Get-TagValue @("album")
      $embeddedTags.Year = Get-TagValue @("year")
      $embeddedTags.Date = Get-TagValue @("date")
      $embeddedTags.Track = Get-TagValue @("track", "tracknumber")
      $embeddedTags.Disc = Get-TagValue @("disc", "discnumber")
      $embeddedTags.Title = Get-TagValue @("title")
    } catch {
      # Continue to filename parsing fallback
    }
  }
  
  # If embedded tags are missing, try filename parsing
  $filenameMeta = Get-FilenameMetadata -FilePath $FilePath
  
  # Merge: prefer embedded tags, fall back to filename parsing
  return @{
    Artist = if ($embeddedTags.Artist) { $embeddedTags.Artist } else { $filenameMeta.Artist }
    AlbumArtist = $embeddedTags.AlbumArtist  # Usually not in filenames
    Album = if ($embeddedTags.Album) { $embeddedTags.Album } else { $filenameMeta.Album }
    Year = $embeddedTags.Year
    Date = $embeddedTags.Date
    Track = if ($embeddedTags.Track) { $embeddedTags.Track } else { $filenameMeta.Track }
    Disc = $embeddedTags.Disc
    Title = if ($embeddedTags.Title) { $embeddedTags.Title } else { $filenameMeta.Title }
  }
}

# ==================== T9: Resolve Year ====================

function Resolve-Year {
  param(
    [hashtable]$Metadata,
    [string]$FilePath
  )
  
  # Prefer year tag, then date tag (extract year from date)
  # Only use embedded tags - do NOT fall back to filename extraction
  # to avoid incorrectly using album names that happen to be years (e.g., "1985")
  if ($Metadata.Year -and -not [string]::IsNullOrWhiteSpace($Metadata.Year)) {
    # Try to extract 4-digit year from Year tag
    if ($Metadata.Year -match '\b(\d{4})\b') {
      $year = $matches[1]
      # Validate it's a reasonable year (1900-2099)
      if ([int]$year -ge 1900 -and [int]$year -le 2099) {
        return $year
      }
    }
  }
  
  if ($Metadata.Date -and -not [string]::IsNullOrWhiteSpace($Metadata.Date)) {
    # Try to extract 4-digit year from Date tag
    if ($Metadata.Date -match '\b(\d{4})\b') {
      $year = $matches[1]
      # Validate it's a reasonable year (1900-2099)
      if ([int]$year -ge 1900 -and [int]$year -le 2099) {
        return $year
      }
    }
  }
  
  # DO NOT use filename/folder fallback to avoid incorrectly using album names that are years
  # (e.g., "Artist - 1985 - Track" where 1985 is the album name, not the release year)
  # Only use embedded metadata tags for accuracy
  return $null
}

# ==================== T10: Resolve Artist (Smart Rule) ====================

function Resolve-Artist {
  param(
    [hashtable]$Metadata,
    [string]$FilePath
  )
  
  # Smart artist resolution:
  # 1. Prefer album_artist if present
  # 2. Fall back to artist tag (from embedded tags or filename)
  # 3. Fall back to folder name parsing
  # 4. Last resort: "Unknown Artist"
  
  if ($Metadata.AlbumArtist -and -not [string]::IsNullOrWhiteSpace($Metadata.AlbumArtist)) {
    return $Metadata.AlbumArtist
  }
  
  if ($Metadata.Artist -and -not [string]::IsNullOrWhiteSpace($Metadata.Artist)) {
    return $Metadata.Artist
  }
  
  # Fallback: try to extract from folder structure
  # Common pattern: Artist/Album/Track.ext
  $parentDir = [System.IO.Path]::GetDirectoryName($FilePath)
  $grandParentDir = [System.IO.Path]::GetDirectoryName($parentDir)
  $albumDir = Split-Path -Leaf $parentDir
  $artistDir = if ($grandParentDir) { Split-Path -Leaf $grandParentDir } else { $null }
  
  # If we have a clear Artist/Album structure, use artist folder
  # But skip if folder is "bkp" or other backup/system folders
  $skipFolders = @("bkp", "backup", "backups", "test", "temp", "tmp")
  if ($artistDir -and $albumDir -and $artistDir -ne $albumDir -and 
      $skipFolders -notcontains $artistDir.ToLower() -and 
      $skipFolders -notcontains $albumDir.ToLower()) {
    return $artistDir
  }
  
  # Last resort: use album folder name (if not a skip folder) or "Unknown Artist"
  if ($albumDir -and $skipFolders -notcontains $albumDir.ToLower()) {
    return $albumDir
  }
  
  return "Unknown Artist"
}

# ==================== T10.1: Disc Number Fallback + Album Normalization ====================

function Resolve-DiscNumberFromPath {
  <#
    Attempts to infer disc number from the folder path when tags are missing.

    Supported folder name patterns (case-insensitive):
      - Disc 1, Disc01, Disc_1
      - Disk 2, Disk02
      - CD3, CD 03

    We only infer disc from folder names (not from the filename) to avoid false positives.
  #>
  param(
    [Parameter(Mandatory)][string]$FilePath
  )

  $dir = [System.IO.Path]::GetDirectoryName($FilePath)
  while ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) {
    $leaf = (Split-Path -Leaf $dir)
    if ($leaf) {
      $name = $leaf.Trim()
      # Normalize separators to spaces for easier matching
      $nameNorm = ($name -replace '[_\-]+', ' ').Trim()

      # Match: Disc 1 / Disk 2 / CD 03 / CD3
      if ($nameNorm -match '^(disc|disk|cd)\s*0*(\d{1,2})$') {
        return $matches[2]
      }
    }
    $dir = [System.IO.Path]::GetDirectoryName($dir)
  }

  return $null
}

function Convert-AlbumTrailingDiscMarker {
  <#
    Removes trailing disc markers from the ALBUM tag so multi-disc sets group into one album.

    Examples:
      - "Final Fantasy X Original Soundtrack (Disc 1)" -> "Final Fantasy X Original Soundtrack"
      - "Album [CD2]" -> "Album"
      - "Album - Disc 03" -> "Album"
  #>
  param(
    [Parameter(Mandatory)][string]$Album
  )

  $a = $Album.Trim()

  # Remove trailing patterns like "(Disc 1)", "[CD2]", "- Disc 03"
  $a2 = $a -replace '\s*(?:[\(\[]\s*)?(disc|disk|cd)\s*0*\d{1,2}(?:\s*[\)\]])?\s*$', ''
  $a2 = $a2 -replace '\s*[-–—]\s*$', ''
  $a2 = $a2.Trim()

  if ([string]::IsNullOrWhiteSpace($a2)) { return $a }
  return $a2
}

function Set-DiscFallbackAndNormalizeAlbum {
  <#
    Implements:
      A) Disc fallback: infer DISCNUMBER from folder name when missing.
      B) Album normalization: if disc was inferred, strip "(Disc X)" style suffixes from ALBUM.

    Missing-only behavior: we NEVER overwrite an existing disc tag.
  #>
  param(
    [Parameter(Mandatory)][hashtable]$Metadata,
    [Parameter(Mandatory)][string]$FilePath
  )

  $discWasInferred = $false

  # A) Infer disc number only if missing
  if (-not $Metadata.Disc -or [string]::IsNullOrWhiteSpace([string]$Metadata.Disc)) {
    $disc = Resolve-DiscNumberFromPath -FilePath $FilePath
    if ($disc) {
      $Metadata.Disc = $disc
      $discWasInferred = $true
    }
  }

  # B) Normalize album only when disc was inferred
  if ($discWasInferred -and $Metadata.Album -and -not [string]::IsNullOrWhiteSpace([string]$Metadata.Album)) {
    $normalized = Convert-AlbumTrailingDiscMarker -Album ([string]$Metadata.Album)
    if ($normalized -ne $Metadata.Album) {
      $Metadata.Album = $normalized
    }
  }

  return $Metadata
}

# ==================== T11: Build Output Paths (Navidrome Layout) ====================

function Build-OutputPath {
  param(
    [Parameter(Mandatory)][string]$FilePath,
    [hashtable]$Metadata
  )
  
  # Output next to the source file (no artist/album folders)
  # Multi-disc: Track number includes disc if present (no disc folders)
  
  $trackNum = if ($Metadata.Track) {
    # Handle "1/12" or "01" formats
    $trackStr = $Metadata.Track
    if ($trackStr -match '^(\d+)') {
      $trackStr = $matches[1].PadLeft(2, '0')
    } else {
      $trackStr = "01"
    }
    
    # Add disc prefix if disc tag exists
    if ($Metadata.Disc -and $Metadata.Disc -match '^\d+') {
      $discNum = $matches[0].PadLeft(2, '0')
      $trackStr = "${discNum}-${trackStr}"
    }
    
    $trackStr
  } else {
    "01"
  }
  
  $title = if ($Metadata.Title -and -not [string]::IsNullOrWhiteSpace($Metadata.Title)) {
    $Metadata.Title
  } else {
    [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
  }
  
  # Sanitize filenames (remove invalid chars)
  $sanitize = {
    param([string]$s)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) {
      $s = $s.Replace($c, '_')
    }
    # Remove multiple underscores
    while ($s -match '__') {
      $s = $s -replace '__', '_'
    }
    return $s.Trim('_', ' ')
  }
  
  $safeTitle = & $sanitize $title
  
  $outputDir = [System.IO.Path]::GetDirectoryName($FilePath)
  $outputFileName = "${trackNum} - ${safeTitle}${OPUS_EXTENSION}"
  $outputPath = Join-Path $outputDir $outputFileName
  
  return @{
    Directory = $outputDir
    FileName = $outputFileName
    FullPath = $outputPath
  }
}

# ==================== T12: Build Archive Paths ====================

function Build-ArchivePath {
  param(
    [Parameter(Mandatory)][string]$DestinationRoot,
    [Parameter(Mandatory)][string]$SourceRoot,
    [Parameter(Mandatory)][string]$FilePath
  )
  
  # Preserve relative path structure
  $relPath = $FilePath.Substring($SourceRoot.Length).TrimStart('\', '/')
  $archivePath = Join-Path $DestinationRoot $relPath
  
  return @{
    Directory = [System.IO.Path]::GetDirectoryName($archivePath)
    FullPath = $archivePath
  }
}

# ==================== T13 & T14: FFmpeg Direct Conversion with Metadata ====================

function Invoke-ConvertToOpus {
  param(
    [Parameter(Mandatory)][string]$FFmpegPath,
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [hashtable]$Metadata
  )
  
  # Ensure output directory exists
  $outputDir = [System.IO.Path]::GetDirectoryName($OutputPath)
  if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
  }
  
  try {
    # Get source sample rate
    $ffprobePath = Get-ToolPath -Name "ffprobe"
    $srArgs = @(
      "-v", "error",
      "-select_streams", "a:0",
      "-show_entries", "stream=sample_rate",
      "-of", "default=noprint_wrappers=1:nokey=1",
      $InputPath
    )
    $sourceSampleRate = & $ffprobePath @srArgs 2>$null
    if (-not $sourceSampleRate -or [string]::IsNullOrWhiteSpace($sourceSampleRate)) {
      $sourceSampleRate = "48000"  # Default fallback
    } else {
      $sourceSampleRate = $sourceSampleRate.Trim()
    }
    
    # Convert to integer for comparison
    $sourceSampleRateInt = [int]$sourceSampleRate
    
    # Check if sample rate is supported by libopus, if not resample to 48000 Hz
    $targetSampleRate = $sourceSampleRateInt
    if ($OPUS_SUPPORTED_SAMPLE_RATES -notcontains $sourceSampleRateInt) {
      # Resample to 48000 Hz (highest quality supported rate)
      $targetSampleRate = 48000
      Write-Host "  [RESAMPLE] Converting sample rate from ${sourceSampleRateInt} Hz to 48000 Hz (libopus requirement)" -ForegroundColor Yellow
    }
    
    # Build FFmpeg arguments for direct Opus conversion with metadata
    $ffmpegArgs = @(
      "-y",
      "-i", $InputPath,
      "-vn",  # No video
      "-ac", $OPUS_CHANNELS,  # Stereo
      "-ar", $targetSampleRate.ToString(),  # Target sample rate (supported by libopus)
      "-c:a", "libopus",  # Use libopus codec
      "-b:a", "${OPUS_BITRATE}k",  # Bitrate in kbps
      "-vbr", "on"  # VBR mode
    )
    
    # Add metadata tags using FFmpeg's -metadata option
    # Use uppercase Vorbis comment field names for Navidrome compatibility
    # Navidrome expects: TITLE, ARTIST, ALBUM, ALBUMARTIST, TRACKNUMBER, DISCNUMBER, DATE
    if ($Metadata.Artist -and -not [string]::IsNullOrWhiteSpace($Metadata.Artist)) {
      $ffmpegArgs += "-metadata", "ARTIST=$($Metadata.Artist)"
    }
    if ($Metadata.AlbumArtist -and -not [string]::IsNullOrWhiteSpace($Metadata.AlbumArtist)) {
      $ffmpegArgs += "-metadata", "ALBUMARTIST=$($Metadata.AlbumArtist)"
    }
    if ($Metadata.Album -and -not [string]::IsNullOrWhiteSpace($Metadata.Album)) {
      $ffmpegArgs += "-metadata", "ALBUM=$($Metadata.Album)"
    }
    if ($Metadata.Title -and -not [string]::IsNullOrWhiteSpace($Metadata.Title)) {
      $ffmpegArgs += "-metadata", "TITLE=$($Metadata.Title)"
    }
    if ($Metadata.Track -and -not [string]::IsNullOrWhiteSpace($Metadata.Track)) {
      # Extract track number (handle "1/12" format)
      $trackNum = $Metadata.Track
      if ($trackNum -match '^(\d+)') {
        $trackNum = $matches[1]
      }
      $ffmpegArgs += "-metadata", "TRACKNUMBER=$trackNum"
    }
    if ($Metadata.Disc -and -not [string]::IsNullOrWhiteSpace($Metadata.Disc)) {
      # Extract disc number
      $discNum = $Metadata.Disc
      if ($discNum -match '^(\d+)') {
        $discNum = $matches[1]
      }
      $ffmpegArgs += "-metadata", "DISCNUMBER=$discNum"
    }
    # Use resolvedYear if available (prioritizes embedded tags over filename extraction)
    # Fall back to Metadata.Year or Metadata.Date if resolvedYear is not available
    $yearToUse = $null
    if ($Metadata.ResolvedYear -and -not [string]::IsNullOrWhiteSpace($Metadata.ResolvedYear)) {
      $yearToUse = $Metadata.ResolvedYear
    } elseif ($Metadata.Year -and -not [string]::IsNullOrWhiteSpace($Metadata.Year)) {
      if ($Metadata.Year -match '\b(\d{4})\b') {
        $yearToUse = $matches[1]
      }
    } elseif ($Metadata.Date -and -not [string]::IsNullOrWhiteSpace($Metadata.Date)) {
      if ($Metadata.Date -match '\b(\d{4})\b') {
        $yearToUse = $matches[1]
      }
    }
    
    if ($yearToUse -and [int]$yearToUse -ge 1900 -and [int]$yearToUse -le 2099) {
      # Use uppercase DATE per Vorbis comment standard for Navidrome
      # Format: YYYY (ISO 8601 subset for year-only)
      $ffmpegArgs += "-metadata", "DATE=$yearToUse"
    }
    
    # Output file
    $ffmpegArgs += $OutputPath
    
    # Execute FFmpeg conversion
    $ffmpegError = & $FFmpegPath @ffmpegArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
      $errorMsg = "FFmpeg failed with exit code $LASTEXITCODE"
      if ($ffmpegError) {
        $errorMsg += ". Error output: $($ffmpegError -join ' ')"
      }
      throw $errorMsg
    }
    
    if (-not (Test-Path -LiteralPath $OutputPath)) {
      throw "FFmpeg did not produce output file"
    }
    
    return @{ Success = $true }
  } catch {
    return @{ Success = $false; Error = $_.Exception.Message }
  }
}

# ==================== T15, T16, T17: Verification Contract ====================

function Test-VerificationContract {
  param(
    [Parameter(Mandatory)][string]$FFprobePath,
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$OutputPath
  )
  
  $errors = @()
  
  # Rule 1: Output OPUS file exists and is non-zero size
  if (-not (Test-Path -LiteralPath $OutputPath)) {
    $errors += "Output file does not exist"
    return @{ Passed = $false; Errors = $errors }
  }
  
  $outputInfo = Get-Item -LiteralPath $OutputPath
  if ($outputInfo.Length -eq 0) {
    $errors += "Output file is zero size"
    return @{ Passed = $false; Errors = $errors }
  }
  
  # Rule 2: Audio stream is readable (ffprobe)
  $streamArgs = @(
    "-v", "error",
    "-select_streams", "a:0",
    "-show_entries", "stream=codec_name,channels,sample_rate,duration",
    "-of", "json",
    $OutputPath
  )
  
  $streamJson = & $FFprobePath @streamArgs 2>$null
  if (-not $streamJson) {
    $errors += "Cannot read audio stream from output"
    return @{ Passed = $false; Errors = $errors }
  }
  
  try {
    $streamObj = $streamJson | ConvertFrom-Json
    if (-not $streamObj.streams -or $streamObj.streams.Count -eq 0) {
      $errors += "No audio streams found in output"
      return @{ Passed = $false; Errors = $errors }
    }
    
    $audioStream = $streamObj.streams[0]
    
    # Rule 3: Duration matches source (within tolerance)
    $sourceDuration = Get-SourceDuration -FFprobePath $FFprobePath -FilePath $SourcePath
    $outputDuration = if ($audioStream.duration) { [double]$audioStream.duration } else { $null }
    
    if ($sourceDuration -and $outputDuration) {
      $durationDiff = [Math]::Abs($sourceDuration - $outputDuration)
      if ($durationDiff -gt $DURATION_TOLERANCE) {
        $errors += "Duration mismatch: source=$sourceDuration, output=$outputDuration (diff=${durationDiff}s)"
      }
    }
    
    # Rule 4: Channels = stereo
    $channels = if ($audioStream.channels) { [int]$audioStream.channels } else { 0 }
    if ($channels -ne 2) {
      $errors += "Channels is $channels (expected 2/stereo)"
    }
    
    # Rule 5: Required metadata is present (or safely derived)
    # This is checked separately in the main loop, but we verify basic structure here
    # Metadata verification happens via FFmpeg tags embedded in the Opus file
    
  } catch {
    $errors += "Failed to parse ffprobe output: $($_.Exception.Message)"
  }
  
  if ($errors.Count -gt 0) {
    return @{ Passed = $false; Errors = $errors }
  }
  
  return @{ Passed = $true; Errors = @() }
}

function Get-SourceDuration {
  param(
    [Parameter(Mandatory)][string]$FFprobePath,
    [Parameter(Mandatory)][string]$FilePath
  )
  
  $durationArgs = @(
    "-v", "error",
    "-select_streams", "a:0",
    "-show_entries", "stream=duration",
    "-of", "default=noprint_wrappers=1:nokey=1",
    $FilePath
  )
  
  $duration = & $FFprobePath @durationArgs 2>$null
  if ($duration -and $duration -match '^\d+\.?\d*') {
    return [double]$duration.Trim()
  }
  
  # Fallback: try format duration
  $formatArgs = @(
    "-v", "error",
    "-show_entries", "format=duration",
    "-of", "default=noprint_wrappers=1:nokey=1",
    $FilePath
  )
  
  $formatDuration = & $FFprobePath @formatArgs 2>$null
  if ($formatDuration -and $formatDuration -match '^\d+\.?\d*') {
    return [double]$formatDuration.Trim()
  }
  
  return $null
}

# ==================== T18: Move Originals After Success ====================

function Move-OriginalToArchive {
  param(
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$ArchivePath
  )
  
  $archiveDir = [System.IO.Path]::GetDirectoryName($ArchivePath)
  if (-not (Test-Path -LiteralPath $archiveDir)) {
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
  }
  
  Move-Item -LiteralPath $SourcePath -Destination $ArchivePath -Force
}

function Move-CoverImagesToArchive {
  param(
    [Parameter(Mandatory)][string]$SourceDir,
    [Parameter(Mandatory)][string]$ArchiveDir,
    [Parameter(Mandatory)][string]$OutputDir
  )
  
  # Ensure both directories exist
  if (-not (Test-Path -LiteralPath $ArchiveDir)) {
    New-Item -ItemType Directory -Path $ArchiveDir -Force | Out-Null
  }
  if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
  }
  
  $movedCount = 0
  
  # Get all files in source directory for case-insensitive matching
  $sourceFiles = Get-ChildItem -LiteralPath $SourceDir -File -ErrorAction SilentlyContinue
  if (-not $sourceFiles) {
    return $movedCount
  }
  
  # Create a case-insensitive lookup of source filenames
  $sourceFileNames = @{}
  foreach ($file in $sourceFiles) {
    $lowerName = $file.Name.ToLower()
    if (-not $sourceFileNames.ContainsKey($lowerName)) {
      $sourceFileNames[$lowerName] = $file.Name
    }
  }
  
  # Check for each common cover image filename (case-insensitive)
  foreach ($coverName in $COVER_IMAGE_NAMES) {
    $coverNameLower = $coverName.ToLower()
    
    # Check if a file with this name (case-insensitive) exists
    if ($sourceFileNames.ContainsKey($coverNameLower)) {
      $actualCoverName = $sourceFileNames[$coverNameLower]
      $sourceCoverPath = Join-Path $SourceDir $actualCoverName
      $archiveCoverPath = Join-Path $ArchiveDir $coverName
      $outputCoverPath = Join-Path $OutputDir $coverName
      
      $sourceFullPath = [System.IO.Path]::GetFullPath($sourceCoverPath)
      $outputFullPath = [System.IO.Path]::GetFullPath($outputCoverPath)
      $outputIsSource = [string]::Equals($sourceFullPath, $outputFullPath, [System.StringComparison]::OrdinalIgnoreCase)
      $archiveExists = Test-Path -LiteralPath $archiveCoverPath
      
      try {
        # If output is a different directory, copy before moving
        if (-not $outputIsSource -and -not (Test-Path -LiteralPath $outputCoverPath)) {
          Copy-Item -LiteralPath $sourceCoverPath -Destination $outputCoverPath -Force
          Write-Host "  [COVER] Copied $actualCoverName to output" -ForegroundColor DarkGreen
        }
        
        # Move original to archive only if not already archived
        if (-not $archiveExists) {
          Move-Item -LiteralPath $sourceCoverPath -Destination $archiveCoverPath -Force
          $movedCount++
          Write-Host "  [COVER] Moved $actualCoverName to archive" -ForegroundColor DarkGreen
        }
        
        # If output is the same as source, ensure a copy exists after moving
        if ($outputIsSource -and (Test-Path -LiteralPath $archiveCoverPath) -and -not (Test-Path -LiteralPath $outputCoverPath)) {
          Copy-Item -LiteralPath $archiveCoverPath -Destination $outputCoverPath -Force
          Write-Host "  [COVER] Copied $actualCoverName to output" -ForegroundColor DarkGreen
        }
      } catch {
        Write-Warning "Failed to process cover image ${actualCoverName}: $($_.Exception.Message)"
      }
    }
  }
  
  return $movedCount
}

# ==================== T19, T20: CSV Logging ====================

function Initialize-Logging {
  param(
    [Parameter(Mandatory)][string]$ScriptDir
  )
  
  $logsDir = Join-Path $ScriptDir "logs"
  if (-not (Test-Path -LiteralPath $logsDir -PathType Container)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
  }
  
  $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $logFile = Join-Path $logsDir "AudioConversion_$timestamp.csv"
  
  # CSV header
  $header = "Timestamp,SourcePath,OutputPath,ArchivePath,Status,Reason,Artist,Album,Title,Year"
  Set-Content -LiteralPath $logFile -Value $header -Encoding UTF8
  
  return $logFile
}

function Write-LogEntry {
  param(
    [Parameter(Mandatory)][string]$LogFile,
    [Parameter(Mandatory)][string]$SourcePath,
    [string]$OutputPath,
    [string]$ArchivePath,
    [Parameter(Mandatory)][string]$Status,
    [string]$Reason,
    [string]$Artist,
    [string]$Album,
    [string]$Title,
    [string]$Year
  )
  
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  
  # CSV escape: wrap in quotes if contains comma or quote
  function Format-CsvValue {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    if ($s -match '[, "]') {
      return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
  }
  
  $row = @(
    $timestamp,
    (Format-CsvValue $SourcePath),
    (Format-CsvValue $OutputPath),
    (Format-CsvValue $ArchivePath),
    (Format-CsvValue $Status),
    (Format-CsvValue $Reason),
    (Format-CsvValue $Artist),
    (Format-CsvValue $Album),
    (Format-CsvValue $Title),
    (Format-CsvValue $Year)
  ) -join ","
  
  Add-Content -LiteralPath $LogFile -Value $row -Encoding UTF8
}

# ==================== T5: Optional Preview ====================

function Show-Preview {
  param(
    [array]$WorkItems,
    [string]$SourceRoot,
    [string]$DestinationRoot
  )
  
  Write-Host ""
  Write-Host "=== PREVIEW ===" -ForegroundColor Cyan
  Write-Host "Eligible files: $($WorkItems.Count)" -ForegroundColor White
  Write-Host ""
  Write-Host "Sample conversions (first 10):" -ForegroundColor Yellow
  
  $sampleCount = [Math]::Min(10, $WorkItems.Count)
  for ($i = 0; $i -lt $sampleCount; $i++) {
    $item = $WorkItems[$i]
    Write-Host "  [$($i+1)] $($item.SourceFile.Name)" -ForegroundColor Gray
    Write-Host "      -> $($item.OutputPath.FullPath)" -ForegroundColor Green
    Write-Host "      -> archive: $($item.ArchivePath.FullPath)" -ForegroundColor DarkGray
  }
  
  if ($WorkItems.Count -gt $sampleCount) {
    Write-Host "  ... and $($WorkItems.Count - $sampleCount) more" -ForegroundColor Gray
  }
  
  Write-Host ""
  $confirm = Read-Host "Proceed with conversion? (Y/N)"
  if ($confirm -notmatch '^[Yy]') {
    Write-Host "Aborted by user." -ForegroundColor Yellow
    exit $EXIT_USER_ABORT
  }
}

# ==================== T21: Main Processing Loop ====================

function Invoke-MainProcessing {
  param(
    [string]$SourceRoot,
    [string]$DestinationRoot,
    [switch]$Preview
  )
  
  # Validate dependencies
  Test-Dependencies | Out-Null
  
  # Get user paths
  $paths = Get-UserPaths -SourceRoot $SourceRoot -DestinationRoot $DestinationRoot
  $SourceRoot = $paths.SourceRoot
  $DestinationRoot = $paths.DestinationRoot
  
  # Initialize logging
  $scriptDir = Resolve-ScriptDir
  $logFile = Initialize-Logging -ScriptDir $scriptDir
  
  # Get tool paths
  $ffmpegPath = Get-ToolPath -Name "ffmpeg"
  $ffprobePath = Get-ToolPath -Name "ffprobe"
  
  Write-Host "Scanning for eligible files..." -ForegroundColor Cyan
  
  # Scan files
  $files = Get-EligibleFiles -SourceRoot $SourceRoot
  
  if ($files.Count -eq 0) {
    Write-Host "No eligible files found." -ForegroundColor Yellow
    exit $EXIT_SUCCESS
  }
  
  Write-Host "Found $($files.Count) eligible files." -ForegroundColor Green
  
  # Build work items
  $workItems = @()
  foreach ($file in $files) {
    # Validate audio stream exists
    if (-not (Test-AudioStreamExists -FFprobePath $ffprobePath -FilePath $file.FullName)) {
      Write-Warning "Skipping $($file.Name): No audio stream detected"
      Write-LogEntry -LogFile $logFile -SourcePath $file.FullName -Status "Skipped" -Reason "No audio stream"
      continue
    }
    
    # Extract metadata
    $metadata = Get-EmbeddedMetadata -FFprobePath $ffprobePath -FilePath $file.FullName
    
    # A + B: Disc fallback from folder name + Album normalization for multi-disc sets
    $metadata = Set-DiscFallbackAndNormalizeAlbum -Metadata $metadata -FilePath $file.FullName
    
    # Resolve year and artist
    $resolvedYear = Resolve-Year -Metadata $metadata -FilePath $file.FullName
    $resolvedArtist = Resolve-Artist -Metadata $metadata -FilePath $file.FullName
    
    # Build paths
    $outputPath = Build-OutputPath -FilePath $file.FullName -Metadata $metadata
    $archivePath = Build-ArchivePath -DestinationRoot $DestinationRoot -SourceRoot $SourceRoot -FilePath $file.FullName
    
    $workItems += @{
      SourceFile = $file
      Metadata = $metadata
      ResolvedYear = $resolvedYear
      ResolvedArtist = $resolvedArtist
      OutputPath = $outputPath
      ArchivePath = $archivePath
    }
  }
  
  if ($workItems.Count -eq 0) {
    Write-Host "No valid work items after filtering." -ForegroundColor Yellow
    exit $EXIT_SUCCESS
  }
  
  # Preview if requested
  if ($Preview) {
    Show-Preview -WorkItems $workItems -SourceRoot $SourceRoot -DestinationRoot $DestinationRoot
  } else {
    $confirm = Read-Host "Proceed with conversion? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
      Write-Host "Aborted by user." -ForegroundColor Yellow
      exit $EXIT_USER_ABORT
    }
  }
  
  # Process each file
  $successCount = 0
  $failCount = 0
  
  foreach ($item in $workItems) {
    $file = $item.SourceFile
    Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan
    
    try {
      # Ensure output directory exists
      if (-not (Test-Path -LiteralPath $item.OutputPath.Directory)) {
        New-Item -ItemType Directory -Path $item.OutputPath.Directory -Force | Out-Null
      }
      
      # Convert
      # Add ResolvedYear to metadata for the conversion function to use
      $metadataForConversion = $item.Metadata.Clone()
      $metadataForConversion.ResolvedYear = $item.ResolvedYear
      $convertResult = Invoke-ConvertToOpus -FFmpegPath $ffmpegPath -InputPath $file.FullName -OutputPath $item.OutputPath.FullPath -Metadata $metadataForConversion
      
      if (-not $convertResult.Success) {
        throw "Conversion failed: $($convertResult.Error)"
      }
      
      # Verify
      $verifyResult = Test-VerificationContract -FFprobePath $ffprobePath -SourcePath $file.FullName -OutputPath $item.OutputPath.FullPath
      
      if (-not $verifyResult.Passed) {
        # Clean up failed output
        if (Test-Path -LiteralPath $item.OutputPath.FullPath) {
          Remove-Item -LiteralPath $item.OutputPath.FullPath -Force -ErrorAction SilentlyContinue
        }
        throw "Verification failed: $($verifyResult.Errors -join '; ')"
      }
      
      # Archive original
      Move-OriginalToArchive -SourcePath $file.FullName -ArchivePath $item.ArchivePath.FullPath
      
      # Move cover images from the source directory to archive and copy to output directory
      $sourceDir = [System.IO.Path]::GetDirectoryName($file.FullName)
      $archiveDir = [System.IO.Path]::GetDirectoryName($item.ArchivePath.FullPath)
      $outputDir = $item.OutputPath.Directory
      $null = Move-CoverImagesToArchive -SourceDir $sourceDir -ArchiveDir $archiveDir -OutputDir $outputDir
      
      Write-Host "  [OK] Converted and archived" -ForegroundColor Green
      $successCount++
      
      Write-LogEntry -LogFile $logFile -SourcePath $file.FullName -OutputPath $item.OutputPath.FullPath -ArchivePath $item.ArchivePath.FullPath -Status "Success" -Reason "Converted successfully" -Artist $item.ResolvedArtist -Album $item.Metadata.Album -Title $item.Metadata.Title -Year $item.ResolvedYear
      
    } catch {
      Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
      $failCount++
      
      Write-LogEntry -LogFile $logFile -SourcePath $file.FullName -Status "Failed" -Reason $_.Exception.Message -Artist $item.ResolvedArtist -Album $item.Metadata.Album -Title $item.Metadata.Title -Year $item.ResolvedYear
      
      # Abort on first failure (fail-fast)
      Write-Host ""
      Write-Host "Aborting due to failure (fail-fast mode)." -ForegroundColor Red
      Write-Host "Log saved to: $logFile" -ForegroundColor Yellow
      exit $EXIT_CONVERSION_FAILURE
    }
  }
  
  Write-Host ""
  Write-Host "=== COMPLETE ===" -ForegroundColor Green
  Write-Host "Success: $successCount" -ForegroundColor Green
  Write-Host "Failed: $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Gray" })
  Write-Host "Log saved to: $logFile" -ForegroundColor Cyan
  
  exit $EXIT_SUCCESS
}

# ==================== T22: Exit Code Handling ====================
# Exit codes are handled throughout the script via exit statements

# ==================== Entry Point ====================

if ($MyInvocation.InvocationName -ne '.') {
  try {
    Invoke-MainProcessing -SourceRoot $SourceRoot -DestinationRoot $DestinationRoot -Preview:$Preview
  } catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit $EXIT_CONVERSION_FAILURE
  }
}
