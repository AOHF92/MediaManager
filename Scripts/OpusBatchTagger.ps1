<#
OpusBatchTagger.ps1
Batch edit Opus (.opus) metadata using kid3-cli (in-place), with optional embedded cover art.

Requires:
- kid3-cli in PATH (recommended) for tag writes
- ffmpeg.exe and ffprobe.exe in PATH (cover art + tag reads)
- Works on Windows 10/11
Notes:
- Opus tags are Vorbis Comments (FIELD=Value).
- kid3-cli edits OPUS Vorbis comments in-place without re-encoding.
- Cover art: embedded as attachment stream in OPUS container (requires FFmpeg).

Safety:
- Default is DRY RUN.
- Use -Commit to write changes (in-place via temp file replacement).
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$RootPath = "",

  [Parameter(Mandatory=$false)]
  [string]$Kid3CliPath = "kid3-cli",

  [Parameter(Mandatory=$false)]
  [string]$FFmpegPath = "ffmpeg",

  [Parameter(Mandatory=$false)]
  [string]$FFprobePath = "ffprobe",

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

function Resolve-Kid3Cli {
  param([string]$Path)
  if ($Path -and (Test-Path $Path)) { return (Resolve-Path $Path).Path }
  if (Test-CommandExists $Path) { return $Path }
  throw "kid3-cli not found. Put kid3-cli in PATH or specify -Kid3CliPath 'C:\path\kid3-cli.exe'"
}

function ConvertTo-Kid3Value {
  param([string]$Value)
  if ($null -eq $Value) { return "" }
  # kid3-cli command strings are quoted with single quotes; escape backslashes and single quotes.
  return $Value.Replace("\\", "\\\\").Replace("'", "\\'")
}

function Invoke-Kid3Write {
  param(
    [string]$Kid3Cli,
    [string]$FilePath,
    [string]$TagName,
    [object]$Value,   # string or string[]
    [switch]$DoCommit
  )

  if (-not $DoCommit) { return $true }

  $t = $TagName.ToLower()
  $isEnumerable = ($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])

  if ($t -eq "genre" -and $isEnumerable) {
    $vals = @()
    foreach ($item in $Value) {
      if ($null -eq $item) { continue }
      $s = ("" + $item).Trim()
      if (-not [string]::IsNullOrWhiteSpace($s)) { $vals += $s }
    }
    if ($vals.Count -eq 0) { return $false }

    # Kid3 stores multi-values as indexed fields: genre[0], genre[1], ...
    # Clear a reasonable number of slots first, then set what we need.
    $cmds = @()
    for ($i = 0; $i -lt 10; $i++) {
      $cmds += "set genre[$i] ''"
    }
    for ($i = 0; $i -lt $vals.Count; $i++) {
      $vEsc = ConvertTo-Kid3Value $vals[$i]
      $cmds += "set genre[$i] '$vEsc'"
    }

    $kid3Args = @()
    foreach ($c in $cmds) { $kid3Args += @("-c", $c) }
    $kid3Args += $FilePath

    $out = & $Kid3Cli @kid3Args 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "kid3-cli failed: $out" }
    return $true
  }

  $v1 = ConvertTo-Kid3Value ("" + $Value)
  $cmd = "set $t '$v1'"
  $out = & $Kid3Cli -c $cmd $FilePath 2>&1 | Out-String
  if ($LASTEXITCODE -ne 0) { throw "kid3-cli failed: $out" }
  return $true
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

function Get-TrackNumberFromFileName {
  param([Parameter(Mandatory)][string]$FilePath)

  $base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
  if (-not $base) { return $null }
  $base = $base.Trim()

  # Pattern: "01-12 ..." (disc-track) -> track=12
  if ($base -match '^\s*0*(\d{1,2})\s*-\s*0*(\d{1,3})(?=\D|$)') {
    return [string]([int]$matches[2])
  }

  # Pattern: "12 - ..." or "12_..." or "12. ..." -> track=12
  if ($base -match '^\s*0*(\d{1,3})(?=\s*[-_.]|[\s]+)') {
    return [string]([int]$matches[1])
  }

  return $null
}

function Invoke-OpustagsWrite {
  param(
    [string]$Opustags,
    [string]$FilePath,
    [hashtable]$Assignments,
    [switch]$DoCommit
  )

  if ($Assignments.Count -eq 0) { return $false }
  if (-not $DoCommit) { return $true }

  try {
    $previousPath = $env:Path
    $opustagsPathResolved = $null
    if ($Opustags -and (Test-Path $Opustags)) {
      $opustagsPathResolved = (Resolve-Path $Opustags).Path
    } else {
      $cmd = Get-Command $Opustags -ErrorAction SilentlyContinue
      if ($cmd) { $opustagsPathResolved = $cmd.Path }
    }

    # If opustags was built with MSYS2, make sure its runtime DLLs are on PATH.
    if ($opustagsPathResolved -and ($opustagsPathResolved -match '\\msys64\\')) {
      $msysRoot = ($opustagsPathResolved -split '\\msys64\\')[0] + '\msys64'
      $msysBins = @(
        (Join-Path $msysRoot 'mingw64\bin'),
        (Join-Path $msysRoot 'usr\bin')
      ) -join ';'
      $env:Path = $msysBins + ';' + $env:Path
    }

    # opustags syntax:
    #   --set FIELD=VALUE        (deletes existing FIELD values and adds a single one)
    #   --delete FIELD --add FIELD=VALUE [--add FIELD=VALUE ...]  (multi-value)
    #
    # Multiple tags with the same FIELD are valid Vorbis comments, so we support
    # passing an array value for a key to write multiple values.
    #
    # References: opustags man page (-a/--add, -d/--delete).
    $opustagsArgs = @("--in-place")

    foreach ($k in $Assignments.Keys) {
      $v = $Assignments[$k]
      if ($null -eq $v) { continue }

      # Treat non-string enumerables (e.g., [string[]]) as multi-value tags.
      $isEnumerable = ($v -is [System.Collections.IEnumerable]) -and -not ($v -is [string])
      if ($isEnumerable) {
        $values = @()
        foreach ($item in $v) {
          if ($null -eq $item) { continue }
          $s = ("" + $item).Trim()
          if (-not [string]::IsNullOrWhiteSpace($s)) { $values += $s }
        }

        if ($values.Count -eq 0) { continue }

        # Replace all existing values for this field.
        $opustagsArgs += "--delete"
        $opustagsArgs += "$k"
        foreach ($val in $values) {
          $opustagsArgs += "--add"
          $opustagsArgs += "${k}=$val"
        }
      } else {
        # Single-value write (replace existing)
        $opustagsArgs += "--set"
        $opustagsArgs += "${k}=$v"
      }
    }
    
    $opustagsExe = if ($opustagsPathResolved) { $opustagsPathResolved } else { $Opustags }

    function Invoke-OpustagsOnce {
      param(
        [string]$ExePath,
        [string[]]$BaseArgs,
        [string]$TargetPath
      )
      $invokeArgs = $BaseArgs + @($TargetPath)
      $output = & $ExePath @invokeArgs 2>&1 | Out-String
      return @{
        ExitCode = $LASTEXITCODE
        Output = $output
      }
    }

    function Convert-ToPosixPath {
      param([string]$WinPath)
      if ($WinPath -match '^[A-Za-z]:\\') {
        $drive = $WinPath.Substring(0,1).ToLower()
        $rest = $WinPath.Substring(2) -replace '\\','/'
        $rest = $rest.TrimStart('/')
        return "/$drive/$rest"
      }
      return $WinPath
    }

    function Convert-ToExtendedWinPath {
      param([string]$WinPath)
      if ($WinPath -match '^[A-Za-z]:\\') {
        return "\\?\$WinPath"
      }
      return $WinPath
    }

    $attempts = @()
    $attempts += @{ Label = "win"; Path = $FilePath }

    # Try extended-length path to avoid MAX_PATH issues.
    if ($FilePath.Length -ge 240 -and $FilePath -match '^[A-Za-z]:\\') {
      $attempts += @{ Label = "win_long"; Path = (Convert-ToExtendedWinPath -WinPath $FilePath) }
    }

    $pathLower = if ($opustagsPathResolved) { $opustagsPathResolved.ToLowerInvariant() } else { "" }
    $isMingw = ($pathLower -like '*\msys64\mingw64\*') -or
               ($pathLower -like '*\msys64\ucrt64\*') -or
               ($pathLower -like '*\msys64\clang64\*')
    $shouldTryPosix = $opustagsPathResolved -and ($opustagsPathResolved -match '\\msys64\\') -and (-not $isMingw)
    if ($shouldTryPosix) {
      $attempts += @{ Label = "posix"; Path = (Convert-ToPosixPath -WinPath $FilePath) }
    }

    $failures = @()
    foreach ($attempt in $attempts) {
      $result = Invoke-OpustagsOnce -ExePath $opustagsExe -BaseArgs $opustagsArgs -TargetPath $attempt.Path
      if ($result.ExitCode -eq 0) { return $true }
      $failures += "$($attempt.Label): exit $($result.ExitCode) - $($result.Output)"
    }

    if ($failures.Count -gt 0) {
      throw ("opustags failed. " + ($failures -join " | "))
    }
    
    return $true
  } catch {
    throw
  } finally {
    if ($previousPath) { $env:Path = $previousPath }
  }
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
    
    # ffprobe JSON can return streams as a single object or an array.
    # Force it to be an array so .Count and [0] always work.
    $streams = @()
    if ($streamInfo -and $null -ne $streamInfo.streams) {
      $streams = @($streamInfo.streams)
    }

    if ($streams.Count -gt 0) {
        $stream = $streams[0]
        
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
        
        # Get sample rate (may be string)
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
$kid3 = Resolve-Kid3Cli -Path $Kid3CliPath
$ffmpeg = Resolve-FFmpegTool -Path $FFmpegPath -ToolName "ffmpeg"
$ffprobe = Resolve-FFmpegTool -Path $FFprobePath -ToolName "ffprobe"

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

$deriveTrackFromFilename = $false

if ($tagName.ToUpper() -eq "TRACKNUMBER") {
  Write-Host ""
  $tnMode = Read-Host "TRACKNUMBER input: (1) Set same value for all  (2) Derive from filename prefix"
  if ($tnMode -eq "2") { $deriveTrackFromFilename = $true }
}

if (-not $deriveTrackFromFilename) {
  $newValue = Read-Host "Enter the value to set for $tagName"
} else {
  # No single shared value when deriving per-file
  $newValue = ""
}

# Multi-value support (Navidrome + Vorbis comments):
# For GENRE we can write multiple values as repeated GENRE tags.
# Enter values separated by ',', ';' or '/' (e.g., "Game; Soundtrack; Rock").
$newValueList = $null
if ($tagName.ToUpper() -eq "GENRE") {
  # IMPORTANT: PowerShell can "unwrap" single-item pipeline output into a scalar string.
  # A scalar string does NOT have a .Count property, so always force an array.
  $parts = @(
    ($newValue -split '\s*[,;/]\s*') |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      ForEach-Object { $_.Trim() }
  )

  if ($parts.Count -gt 1) {
    # Preserve order while removing exact duplicates
    $seen = @{}
    $deduped = @()
    foreach ($p in $parts) {
      if (-not $seen.ContainsKey($p)) { $seen[$p] = $true; $deduped += $p }
    }
    $newValueList = $deduped
  }
}

$doEmbed = Read-Host "Embed cover art if found in each album folder? (y/n)"
$embedCover = $doEmbed -match '^(y|yes)$'

# Logging
$logDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $logDir "OpusTagEdits_$stamp.csv"
"File,AlbumFolder,Tag,OldValue,NewValue,Changed,EmbeddedCover,Result" | Out-File -Encoding UTF8 -FilePath $logPath

$files = @(Get-OpusFiles -Path $RootPath)
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
  $valueForLog = $null

  try {
    if ($shouldChange) {
      # Tag writing is done exclusively via kid3-cli (in-place).
      $valueToWrite = $null

      if ($deriveTrackFromFilename -and $tagName.ToUpper() -eq "TRACKNUMBER") {
        $derived = Get-TrackNumberFromFileName -FilePath $filePath
        if (-not $derived) {
          throw "Could not derive TRACKNUMBER from filename: $($f.Name)"
        }
        $valueToWrite = $derived
      } else {
        $valueToWrite = if ($newValueList) { $newValueList } else { $newValue }
      }

      $valueForLog = $valueToWrite
      $didWrite = Invoke-Kid3Write -Kid3Cli $kid3 -FilePath $filePath -TagName $tagName -Value $valueToWrite -DoCommit:$Commit
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
    $skippedCount++
  }

  $changedFlag = $(if ($shouldChange) { "Yes" } else { "No" })
  $coverFlag = $(if ($didCover) { "Yes" } else { "No" })

  # CSV-safe-ish escaping
  $oldEsc = ("" + $old).Replace('"','''')
  $newDisplay = if ($deriveTrackFromFilename -and $tagName.ToUpper() -eq "TRACKNUMBER") {
    ("" + $valueForLog)
  } else {
    if ($newValueList) { ($newValueList -join "; ") } else { $newValue }
  }
  $newEsc = ("" + $newDisplay).Replace('"','''')

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
