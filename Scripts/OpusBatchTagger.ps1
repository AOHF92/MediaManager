<#
OpusBatchTagger.ps1
Batch edit Opus (.opus) metadata using Kid3 CLI (kid3-cli), with optional embedded cover art.

Requires:
- kid3-cli.exe in PATH (required)
- ffprobe.exe in PATH (used to read existing tags)
- ffmpeg.exe in PATH (only needed if you enable cover-art embedding)
- Works on Windows 10/11
Notes:
- Opus tags are Vorbis Comments.
- kid3-cli supports setting common Vorbis comment fields using: kid3-cli -c "set artist 'Name'" file.opus
- Cover art: embedded as an attachment stream in the Opus container (requires FFmpeg).

Safety:
- Default is DRY RUN.
- Use -Commit to write changes (in-place via temp file replacement).
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$RootPath = "",

  [Parameter(Mandatory=$false)]
  [string]$FFmpegPath = "ffmpeg",

  [Parameter(Mandatory=$false)]
  [string]$FFprobePath = "ffprobe",

  [Parameter(Mandatory=$false)]
  [string]$Kid3CliPath = "kid3-cli",

  [Parameter(Mandatory=$false)]
  [switch]$Commit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-CommandExists {
  param([string]$Command)
  $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

function Resolve-FFmpegTool {
  param([string]$Path, [string]$ToolName)
  if ($Path -and (Test-Path $Path)) { return (Resolve-Path $Path).Path }
  if (Test-CommandExists $Path) { return $Path }
  throw "$ToolName not found. Put $ToolName.exe in PATH or specify -${ToolName}Path 'C:\path\$ToolName.exe'"
}

function Resolve-Kid3CliTool {
  param([string]$Path)
  if ($Path -and (Test-Path $Path)) { return (Resolve-Path $Path).Path }
  if (Test-CommandExists $Path) { return $Path }
  throw "kid3-cli not found. Put kid3-cli.exe in PATH or specify -Kid3CliPath 'C:\\path\\kid3-cli.exe'"
}

function Read-FolderPath {
  param([string]$PromptText)
  while ($true) {
    $p = Read-Host $PromptText
    if ([string]::IsNullOrWhiteSpace($p)) { continue }
    if (Test-Path $p) { return (Resolve-Path $p).Path }
    Write-Host "Path not found: $p" -ForegroundColor Yellow
  }
}

function Get-OpusFiles {
  param([string]$Path)
  Get-ChildItem -LiteralPath $Path -Recurse -File -Filter "*.opus" |
    Sort-Object FullName
}

function Find-CoverArtFile {
  param([string]$AlbumFolder)

  $patterns = @(
    "cover.jpg","cover.jpeg","cover.png","folder.jpg","folder.jpeg","folder.png",
    "front.jpg","front.jpeg","front.png","album.jpg","album.jpeg","album.png"
  )

  foreach ($name in $patterns) {
    $candidate = Join-Path $AlbumFolder $name
    if (Test-Path -LiteralPath $candidate) { return $candidate }
  }

  # fallback: any single image file if present
  $img = Get-ChildItem -LiteralPath $AlbumFolder -File |
    Where-Object { $_.Extension -match '^\.(jpg|jpeg|png)$' } |
    Sort-Object Length -Descending |
    Select-Object -First 1

  if ($img) { return $img.FullName }
  return $null
}

function Get-TagValue {
  param(
    [string]$FFprobe,
    [string]$FilePath,
    [string]$TagName
  )

  # Use ffprobe to read format tags (Vorbis comments in OPUS)
  # Map common tag names to ffprobe format tag names
  $tagMap = @{
    "TITLE" = "title"
    "ARTIST" = "artist"
    "ALBUM" = "album"
    "ALBUMARTIST" = "album_artist"
    "TRACKNUMBER" = "track"
    "DISCNUMBER" = "disc"
    "DATE" = "date"
    "GENRE" = "genre"
    "COMMENT" = "comment"
  }

  # Use mapped name if available, otherwise use lowercase of provided tag
  $probeTag = if ($tagMap.ContainsKey($TagName.ToUpper())) {
    $tagMap[$TagName.ToUpper()]
  } else {
    $TagName.ToLower()
  }

  $probeArgs = @(
    "-v", "error",
    "-show_entries", "format_tags=$probeTag",
    "-of", "default=noprint_wrappers=1:nokey=1",
    $FilePath
  )

  $result = & $FFprobe @probeArgs 2>$null
  if ($result -and -not [string]::IsNullOrWhiteSpace($result)) {
    return $result.Trim()
  }
  return $null
}

function Invoke-Kid3Write {
  param(
    [string]$Kid3Cli,
    [string]$FilePath,
    [string]$TagName,
    [string]$Value,
    [switch]$DoCommit
  )

  if ([string]::IsNullOrWhiteSpace($TagName)) { return $false }
  if (-not $DoCommit) { return $true }

  # kid3-cli uses its own command parser. When using single quotes for the value,
  # a single quote inside the value must be escaped with a backslash.
  # Example from the kid3-cli manual: set title 'I\'ll be there for you'
  $escapedValue = ("" + $Value).Replace("\\", "\\\\").Replace("'", "\\'")

  # Map our menu tag names to kid3 field names.
  # (Kid3 uses lowercase vorbis field names for Opus/Vorbis comments.)
  $tagMap = @{
    "TITLE"       = "title"
    "ARTIST"      = "artist"
    "ALBUM"       = "album"
    "ALBUMARTIST" = "albumartist"
    "TRACKNUMBER" = "track"
    "DISCNUMBER"  = "disc"
    "DATE"        = "date"
    "GENRE"       = "genre"
    "COMMENT"     = "comment"
  }

  $kid3Field = if ($tagMap.ContainsKey($TagName.ToUpper())) { $tagMap[$TagName.ToUpper()] } else { $TagName }

  # Build the command as a *single* -c argument.
  $cmd = "set $kid3Field '$escapedValue'"

  $allOutput = & $Kid3Cli -c $cmd $FilePath 2>&1 | Out-String
  if ($LASTEXITCODE -ne 0) {
    throw "kid3-cli failed with code $LASTEXITCODE. Output: $allOutput"
  }

  return $true
}

function Invoke-FFmpegWrite {
  param(
    [string]$FFmpeg,
    [string]$FilePath,
    [hashtable]$Assignments,
    [switch]$DoCommit
  )

  # FFmpeg implementation - NOTE: FFmpeg cannot reliably write OPUS metadata
  # This function is kept as fallback but will likely fail
  if ($Assignments.Count -eq 0) { return $false }

  # Map tag names to FFmpeg metadata format
  $tagMap = @{
    "TITLE" = "title"
    "ARTIST" = "artist"
    "ALBUM" = "album"
    "ALBUMARTIST" = "album_artist"
    "TRACKNUMBER" = "track"
    "DISCNUMBER" = "disc"
    "DATE" = "date"
    "GENRE" = "genre"
    "COMMENT" = "comment"
  }

  $metadataArgs = @()
  foreach ($k in $Assignments.Keys) {
    $v = $Assignments[$k]
    if ($null -eq $v) { continue }
    $metaKey = if ($tagMap.ContainsKey($k.ToUpper())) { $tagMap[$k.ToUpper()] } else { $k.ToLower() }
    $metadataArgs += "-metadata"
    $metadataArgs += "${metaKey}=$v"
  }

  if ($metadataArgs.Count -eq 0) { return $false }
  if (-not $DoCommit) { return $true }

  # Use .opus extension for temp file so FFmpeg recognizes the format
  $tempFile = [System.IO.Path]::ChangeExtension($FilePath, ".temp.opus")
  try {
    # CRITICAL: FFmpeg with -c copy cannot modify OPUS metadata (known limitation)
    # Skip -c copy attempt and go straight to re-encoding with same settings
    Write-Host "  [INFO] Re-encoding OPUS with metadata..." -ForegroundColor Cyan
    
    $ffprobePath = Resolve-FFmpegTool -Path "ffprobe" -ToolName "ffprobe"
    
    # Get original codec settings
    # Try format-level bitrate first (more reliable for OPUS)
    $formatArgs = @("-v", "error", "-show_entries", "format=bit_rate", "-of", "json", $FilePath)
    $formatInfoJson = & $ffprobePath @formatArgs 2>$null
    $formatInfo = $formatInfoJson | ConvertFrom-Json
    
    # Get stream-level settings
    $streamArgs = @("-v", "error", "-select_streams", "a:0", "-show_entries", "stream=sample_rate,channels", "-of", "json", $FilePath)
    $streamInfoJson = & $ffprobePath @streamArgs 2>$null
    $streamInfo = $streamInfoJson | ConvertFrom-Json
    
    if ($streamInfo.streams -and $streamInfo.streams.Count -gt 0) {
        $stream = $streamInfo.streams[0]
        
        # Get bitrate from format (preferred) or default to 128k
        $bitrate = 128
        if ($formatInfo.format) {
          $bitrateProp = $formatInfo.format.PSObject.Properties | Where-Object { $_.Name -eq "bit_rate" }
          if ($bitrateProp) {
            $bitrateValue = $bitrateProp.Value
            if ($bitrateValue -and $bitrateValue -ne "N/A" -and $bitrateValue -ne "0") {
              try {
                $bitrate = [math]::Round([int]$bitrateValue / 1000)
                if ($bitrate -lt 32) { $bitrate = 128 }  # Sanity check
              } catch {
                $bitrate = 128
              }
            }
          }
        }
        
        # Get sample rate
        $sampleRate = 48000
        $srProp = $stream.PSObject.Properties | Where-Object { $_.Name -eq "sample_rate" }
        if ($srProp) {
          $srValue = $srProp.Value
          if ($srValue) {
            try {
              $sampleRate = [int]$srValue
            } catch {
              $sampleRate = 48000
            }
          }
        }
        
        # Get channels
        $channels = 2
        $chProp = $stream.PSObject.Properties | Where-Object { $_.Name -eq "channels" }
        if ($chProp) {
          $chValue = $chProp.Value
          if ($chValue) {
            try {
              $channels = [int]$chValue
            } catch {
              $channels = 2
            }
          }
        }
        
        # Re-encode with same settings + metadata
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        $ffmpegArgs = @(
          "-y", "-i", $FilePath,
          "-c:a", "libopus",
          "-b:a", "${bitrate}k",
          "-ar", $sampleRate.ToString(),
          "-ac", $channels.ToString(),
          "-vbr", "on"
        ) + $metadataArgs + @($tempFile)
        
        $allOutput = & $FFmpeg @ffmpegArgs 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
          throw "FFmpeg re-encode failed with code $LASTEXITCODE. Output: $allOutput"
        }
        
        if (-not (Test-Path -LiteralPath $tempFile)) {
          throw "FFmpeg did not create output file after re-encode"
        }
        
        # Verify metadata was written
        $firstKey = ($Assignments.Keys | Select-Object -First 1)
        $expectedValue = $Assignments[$firstKey]
        $verifyTagName = if ($tagMap.ContainsKey($firstKey.ToUpper())) { $tagMap[$firstKey.ToUpper()] } else { $firstKey.ToLower() }
        $writtenValue = Get-TagValue -FFprobe $ffprobePath -FilePath $tempFile -TagName $verifyTagName
        
      if ($writtenValue -ne $expectedValue) {
        throw "Metadata verification failed: expected '$expectedValue', got '$writtenValue'"
      }
    } else {
      throw "Could not read original codec settings for re-encoding"
    }
    
    Move-Item -LiteralPath $tempFile -Destination $FilePath -Force
    return $true
  } catch {
    if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
    throw
  }
  return $true
}

function Set-CoverArt {
  param(
    [string]$FFmpeg,
    [string]$FilePath,
    [string]$CoverPath,
    [switch]$DoCommit
  )

  if (-not $CoverPath) { return $false }
  if (-not (Test-Path -LiteralPath $CoverPath)) { return $false }

  if (-not $DoCommit) { return $true }

  # Embed cover art as attachment stream in OPUS container
  # -map 0 copies all streams from input, -map 1 adds the cover image
  # -c copy for audio, -c:v:1 copy for the image attachment
  # -disposition:v:1 attached_pic marks it as cover art
  # Use .opus extension for temp file so FFmpeg recognizes the format
  $tempFile = [System.IO.Path]::ChangeExtension($FilePath, ".temp.opus")
  try {
    $ffmpegArgs = @(
      "-y",
      "-i", $FilePath,
      "-i", $CoverPath,
      "-map", "0",  # All streams from audio file
      "-map", "1",  # Cover image
      "-c", "copy",  # Copy audio without re-encoding
      "-c:v:1", "copy",  # Copy image stream
      "-disposition:v:1", "attached_pic",  # Mark as cover art
      $tempFile
    )

    $null = & $FFmpeg @ffmpegArgs 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
      throw "FFmpeg exited with code $LASTEXITCODE"
    }

    if (-not (Test-Path -LiteralPath $tempFile)) {
      throw "FFmpeg did not create output file"
    }

    # Replace original with temp file
    Move-Item -LiteralPath $tempFile -Destination $FilePath -Force
    return $true
  } catch {
    # Clean up temp file on error
    if (Test-Path -LiteralPath $tempFile) {
      Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
    throw
  }
}

# ---------------- Main ----------------

# Resolve required tools
$kid3 = Resolve-Kid3CliTool -Path $Kid3CliPath
$ffprobe = Resolve-FFmpegTool -Path $FFprobePath -ToolName "ffprobe"

# ffmpeg is only needed for cover art embedding (if you say yes later), but we resolve it up-front
$ffmpeg = Resolve-FFmpegTool -Path $FFmpegPath -ToolName "ffmpeg"

if ([string]::IsNullOrWhiteSpace($RootPath)) {
  $RootPath = Read-FolderPath "Enter root folder to process (e.g. D:\Media\Music)"
} else {
  $RootPath = (Resolve-Path $RootPath).Path
}

Write-Host ""
Write-Host "Root: $RootPath"
Write-Host ("Mode: " + ($(if ($Commit) { "COMMIT (in-place)" } else { "DRY RUN (no changes)" })))
Write-Host ""

# Tag menu (you can still type a custom tag)
$commonTags = @(
  "TITLE","ARTIST","ALBUM","ALBUMARTIST","TRACKNUMBER","DISCNUMBER","DATE","GENRE","COMMENT"
)

Write-Host "Common tags:"
for ($i=0; $i -lt $commonTags.Count; $i++) {
  Write-Host ("  [{0}] {1}" -f ($i+1), $commonTags[$i])
}
Write-Host "  [0] Custom tag name"
Write-Host ""

$choice = Read-Host "Choose tag number"
[string]$tagName = ""

if ($choice -eq "0") {
  $tagName = Read-Host "Enter custom tag name (e.g. GENRE, ALBUMARTIST, COMPOSER)"
} else {
  $idx = [int]$choice - 1
  if ($idx -lt 0 -or $idx -ge $commonTags.Count) { throw "Invalid tag selection." }
  $tagName = $commonTags[$idx]
}

if ([string]::IsNullOrWhiteSpace($tagName)) { throw "No tag selected." }

$mode = Read-Host "Apply mode: (1) Fill missing only  (2) Overwrite existing"
$fillMissingOnly = $true
if ($mode -eq "2") { $fillMissingOnly = $false }

$newValue = Read-Host "Enter the value to set for $tagName"

$doEmbed = Read-Host "Embed cover art if found in each album folder? (y/n)"
$embedCover = $doEmbed -match '^(y|yes)$'

# Logging
$logDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $logDir "OpusTagEdits_$stamp.csv"
"File,AlbumFolder,Tag,OldValue,NewValue,Changed,EmbeddedCover,Result" | Out-File -Encoding UTF8 -FilePath $logPath

$files = Get-OpusFiles -Path $RootPath
Write-Host ""
Write-Host ("Found {0} opus files." -f $files.Count)
Write-Host ""

$changedCount = 0
$coverCount = 0
$skippedCount = 0

foreach ($f in $files) {
  $filePath = $f.FullName
  $albumFolder = $f.DirectoryName

  $old = Get-TagValue -FFprobe $ffprobe -FilePath $filePath -TagName $tagName

  $shouldChange = $true
  if ($fillMissingOnly) {
    if ($null -ne $old -and -not [string]::IsNullOrWhiteSpace([string]$old)) {
      $shouldChange = $false
    }
  }

  $didWrite = $false
  $didCover = $false
  $result = "OK"

  try {
    if ($shouldChange) {
      # Write tags using kid3-cli (in-place)
      $didWrite = Invoke-Kid3Write -Kid3Cli $kid3 -FilePath $filePath -TagName $tagName -Value $newValue -DoCommit:$Commit
      if ($didWrite) { $changedCount++ } else { $skippedCount++ }
    } else {
      $skippedCount++
    }

    if ($embedCover) {
      $cover = Find-CoverArtFile -AlbumFolder $albumFolder
      if ($cover) {
        $didCover = Set-CoverArt -FFmpeg $ffmpeg -FilePath $filePath -CoverPath $cover -DoCommit:$Commit
        if ($didCover) { $coverCount++ }
      }
    }
  } catch {
    $result = ("ERROR: " + $_.Exception.Message.Replace('"',''''))
  }

  $changedFlag = $(if ($shouldChange) { "Yes" } else { "No" })
  $coverFlag = $(if ($didCover) { "Yes" } else { "No" })

  # CSV-safe-ish escaping
  $oldEsc = ("" + $old).Replace('"','''')
  $newEsc = ("" + $newValue).Replace('"','''')

  ('"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}"' -f
    $filePath.Replace('"',''''),
    $albumFolder.Replace('"',''''),
    $tagName.Replace('"',''''),
    $oldEsc,
    $newEsc,
    $changedFlag,
    $coverFlag,
    $result.Replace('"','''')
  ) | Out-File -Encoding UTF8 -Append -FilePath $logPath
}

Write-Host ""
Write-Host "Done."
Write-Host ("Changed tag on: {0}" -f $changedCount)
Write-Host ("Embedded cover on: {0}" -f $coverCount)
Write-Host ("Skipped: {0}" -f $skippedCount)
Write-Host ("Log: {0}" -f $logPath)
Write-Host ""
