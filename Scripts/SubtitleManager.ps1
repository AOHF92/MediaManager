#requires -Version 5.1
<#
SubtitleManager.ps1
Non-recursive MKV subtitle integrator (sidecar -> embedded) using MKVToolNix mkvmerge.

Goals (current workflow):
- Focus on subtitle integration (external .srt/.ass/.sup/.vtt into .mkv)
- Non-interactive (no prompts)
- Non-recursive (single folder only)
- Safe swap workflow (temp -> verify -> swap)
- Auto-clean temp files immediately after successful swap
- Deterministic language tagging from filename tokens (EN/ENG, JA/JPN, etc.)

Hard dependency:
- mkvmerge (MKVToolNix) must be installed and available in PATH.

Examples:
  .\SubtitleManager.ps1 -Directory "C:\Downloads\Kaiju"
  .\SubtitleManager.ps1 -Directory "C:\Downloads\Kaiju" -DryRun
  .\SubtitleManager.ps1 -Directory "C:\Downloads\Kaiju" -DefaultLanguage eng
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$Directory,

    # Default subtitle language when no recognizable tag exists in filename
    [ValidateSet('eng','jpn','spa','por','fra','deu','ita','rus','kor','zho')]
    [string]$DefaultLanguage = 'eng',

    # If true, keeps a .bak copy of the original MKV after successful integration
    [switch]$KeepBackup,

    # If false, keeps sidecar subtitle files after successful integration (default: deletes them)
    [switch]$NoDeleteSidecars,

    # If true, prints planned actions without modifying anything
    [switch]$DryRun
)

Set-StrictMode -Version Latest

function Assert-Dependency {
    if (-not (Get-Command mkvmerge -ErrorAction SilentlyContinue)) {
        throw "Dependency missing: 'mkvmerge' not found in PATH. Install MKVToolNix and ensure mkvmerge.exe is accessible."
    }
}

function Test-FileLocked {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    
    if (-not (Test-Path -LiteralPath $FilePath)) {
        return $false
    }
    
    try {
        $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        return $false
    } catch {
        # If we can't open the file, it's likely locked
        return $true
    }
}

function Invoke-MkvMerge {
    param([Parameter(Mandatory)][string[]]$ArgsArray)

    # Use .NET Process class for better control over argument quoting
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'mkvmerge'
    $psi.Arguments = ($ArgsArray | ForEach-Object { 
        if ($_ -match '\s') { "`"$_`"" } else { $_ } 
    }) -join ' '
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $null = $process.Start()
    $process.WaitForExit()
    return $process.ExitCode
}

function Get-ExistingSubtitleTracks {
    param([Parameter(Mandatory = $true)][string]$MkvPath)

    if (-not (Test-Path -LiteralPath $MkvPath)) {
        return @()
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'mkvmerge'
        $psi.Arguments = "--identify `"$MkvPath`""
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $null = $process.Start()
        $output = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        
        if ($process.ExitCode -ne 0) {
            return @()
        }

        $existingTracks = @()
        $lines = $output -split "`r?`n"
        
        foreach ($line in $lines) {
            # mkvmerge --identify output format examples:
            # "Track ID 2: subtitles (S_TEXT/UTF8) [language:jpn]"
            # "Track ID 3: subtitles (S_TEXT/ASS) (language:eng)"
            # Try multiple patterns to catch different formats
            if ($line -match 'Track ID \d+:\s+subtitles') {
                # Try [language:xxx] format
                if ($line -match '\[language:(\w+)\]') {
                    $lang = $Matches[1].ToLower()
                    if ($lang -notin $existingTracks) {
                        $existingTracks += $lang
                    }
                }
                # Try (language:xxx) format
                elseif ($line -match '\(language:(\w+)\)') {
                    $lang = $Matches[1].ToLower()
                    if ($lang -notin $existingTracks) {
                        $existingTracks += $lang
                    }
                }
            }
        }
        
        return $existingTracks
    } catch {
        # If identification fails, return empty array (will attempt to add anyway)
        return @()
    }
}

function Get-LanguageFromFileName {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][string]$Fallback
    )

    $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    $patterns = @(
        @{ Code='eng'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:en|eng|english)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='jpn'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:ja|jpn|japanese)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='spa'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:es|spa|spanish|espanol|español)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='por'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:pt|por|portuguese|portugues|português)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='fra'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:fr|fra|french|francais|français)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='deu'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:de|deu|ger|german|deutsch)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='ita'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:it|ita|italian|italiano)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='rus'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:ru|rus|russian)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='kor'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:ko|kor|korean)(?:$|[.\s_\-\)\]\}\[])' },
        @{ Code='zho'; Rx='(?i)(?:^|[.\s_\-\(\[\{])(?:zh|zho|chi|chinese|mandarin|cantonese)(?:$|[.\s_\-\)\]\}\[])' }
    )

    foreach ($p in $patterns) {
        if ($name -match $p.Rx) { return $p.Code }
    }

    return $Fallback
}

function Get-TrackNameForLang {
    param([Parameter(Mandatory=$true)][string]$LangCode)

    switch ($LangCode) {
        'eng' { return 'English' }
        'jpn' { return 'Japanese' }
        'spa' { return 'Spanish' }
        'por' { return 'Portuguese' }
        'fra' { return 'French' }
        'deu' { return 'German' }
        'ita' { return 'Italian' }
        'rus' { return 'Russian' }
        'kor' { return 'Korean' }
        'zho' { return 'Chinese' }
        default { return $LangCode }
    }
}

function Get-EpisodeNumber {
    param([Parameter(Mandatory = $true)][string]$FileName)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    # Try S01E01, S1E1, etc.
    if ($name -match '(?i)S(\d+)E(\d+)') {
        return [int]$Matches[2]
    }

    # Try E01, E1, Ep 01, Episode 01, etc.
    if ($name -match '(?i)(?:^|[.\s_\-\(\[\{])(?:E|Ep|Episode)\s*(\d+)(?:$|[.\s_\-\)\]\}])') {
        return [int]$Matches[1]
    }

    # Try [01], [1], etc.
    if ($name -match '\[(\d+)\]') {
        return [int]$Matches[1]
    }

    # Try standalone episode number at start or after common separators
    if ($name -match '(?i)(?:^|[.\s_\-\(\[\{])(\d{1,3})(?:$|[.\s_\-\)\]\}])') {
        $epNum = [int]$Matches[1]
        # Only return if it's a reasonable episode number (1-999)
        if ($epNum -ge 1 -and $epNum -le 999) {
            return $epNum
        }
    }

    return $null
}

function Get-MatchingSidecars {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$Video
    )

    $dir = $Video.DirectoryName
    $base = $Video.BaseName
    $videoEpNum = Get-EpisodeNumber -FileName $Video.Name

    $allSidecars = Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.srt','.ass','.sup','.vtt') } |
        Sort-Object Name

    $sidecars = @()
    foreach ($sub in $allSidecars) {
        $subBase = $sub.BaseName
        
        # First try: exact base name match or prefix match
        if ($subBase -eq $base -or $subBase -like "$base.*") {
            $sidecars += $sub
            continue
        }

        # Second try: episode number matching (if we found an episode number in the video)
        if ($null -ne $videoEpNum) {
            $subEpNum = Get-EpisodeNumber -FileName $sub.Name
            if ($null -ne $subEpNum -and $subEpNum -eq $videoEpNum) {
                $sidecars += $sub
            }
        }
    }

    # Ensure we always return an array (even if empty or single item)
    return ,@($sidecars)
}

function Invoke-IntegrateSubtitles {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$Video,
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Sidecars
    )

    if ($Sidecars.Count -eq 0) { return $false }

    $videoPath = $Video.FullName
    $dir = $Video.DirectoryName
    $base = $Video.BaseName

    # Check existing subtitle tracks to avoid duplicates
    $existingTracks = Get-ExistingSubtitleTracks -MkvPath $videoPath
    
    # Filter out sidecars that already exist as tracks (by language code)
    $sidecarsToAdd = @()
    foreach ($s in $Sidecars) {
        $lang = Get-LanguageFromFileName -FileName $s.Name -Fallback $DefaultLanguage
        if ($lang -notin $existingTracks) {
            $sidecarsToAdd += $s
        } else {
            Write-Host ("[SubtitleManager] Skipping '{0}' - subtitle track with language '{1}' already exists in '{2}'" -f $s.Name, $lang, $Video.Name)
        }
    }

    if ($sidecarsToAdd.Count -eq 0) {
        Write-Host ("[SubtitleManager] No new subtitles to add for '{0}' (all already exist)" -f $Video.Name)
        return $false
    }

    $tempPath = Join-Path $dir ("{0}__converting__subs.tmp.mkv" -f $base)
    $bakPath  = Join-Path $dir ("{0}.bak" -f $Video.Name)

    $mkvArgs = @(
        '-o'
        $tempPath
        $videoPath
    )

    foreach ($s in $sidecarsToAdd) {
        $lang = Get-LanguageFromFileName -FileName $s.Name -Fallback $DefaultLanguage
        $tname = Get-TrackNameForLang -LangCode $lang
        $mkvArgs += '--language'
        $mkvArgs += ("0:{0}" -f $lang)
        $mkvArgs += '--track-name'
        $mkvArgs += ("0:{0}" -f $tname)
        $mkvArgs += $s.FullName
    }

    # Create a display version with only filenames (not full paths)
    $displayArgs = @()
    for ($i = 0; $i -lt $mkvArgs.Length; $i++) {
        $arg = $mkvArgs[$i]
        # Check if this looks like a file path (starts with drive letter or UNC path)
        if ($arg -match '^[A-Za-z]:\\' -or $arg -match '^\\\\') {
            # It's a file path, show only filename
            $displayArgs += Split-Path -Leaf $arg
        } else {
            $displayArgs += $arg
        }
    }
    Write-Host ("[SubtitleManager] mkvmerge {0}" -f ($displayArgs -join ' '))

    if ($DryRun) { return $false }

    if (-not $PSCmdlet.ShouldProcess($videoPath, "Integrate {0} sidecar subtitle(s)" -f $sidecarsToAdd.Count)) {
        return $false
    }

    # Use helper function to properly handle argument arrays with spaces
    $exitCode = Invoke-MkvMerge -ArgsArray $mkvArgs

    if ($exitCode -ne 0) {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
        throw "mkvmerge failed (ExitCode=$exitCode) for '$videoPath'."
    }

    if (-not (Test-Path -LiteralPath $tempPath)) {
        throw "mkvmerge did not produce temp output: '$tempPath'."
    }

    # Check if video file is locked before attempting to modify it
    if (Test-FileLocked -FilePath $videoPath) {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
        throw "File is locked: '$videoPath' is being used by another process. Please close any programs using this file and try again."
    }

    try {
        if ($KeepBackup) {
            if (Test-Path -LiteralPath $bakPath) { Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue }
            Move-Item -LiteralPath $videoPath -Destination $bakPath -Force
        } else {
            Remove-Item -LiteralPath $videoPath -Force
        }

        Move-Item -LiteralPath $tempPath -Destination $videoPath -Force
    } catch {
        # Clean up temp file if move failed
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
        
        # Check if the error is due to file being locked
        if ($_.Exception.Message -match 'being used by another process' -or $_.Exception.Message -match 'cannot access') {
            throw "File is locked: '$videoPath' is being used by another process. Please close any programs using this file and try again."
        }
        throw
    }

    if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }

    # Delete sidecar files by default (unless -NoDeleteSidecars is specified)
    if (-not $NoDeleteSidecars) {
        foreach ($s in $sidecarsToAdd) {
            if (Test-Path -LiteralPath $s.FullName) {
                Remove-Item -LiteralPath $s.FullName -Force -ErrorAction SilentlyContinue
                Write-Host ("[SubtitleManager] Deleted sidecar: '{0}'" -f $s.Name)
            }
        }
    }

    return $true
}

Assert-Dependency

$videos = Get-ChildItem -LiteralPath $Directory -File -Filter *.mkv | Sort-Object Name

$changed = 0
foreach ($v in $videos) {
    if ($v.Name -like '*__converting__*' -or $v.Name -like '*.tmp.mkv' -or $v.Name -like '*.subclean.tmp.mkv') {
        continue
    }

    $subs = Get-MatchingSidecars -Video $v
    if ($null -eq $subs -or @($subs).Count -eq 0) { continue }

    try {
        $did = Invoke-IntegrateSubtitles -Video $v -Sidecars $subs
        if ($did) { $changed++ }
    } catch {
        $errorMsg = $_.Exception.Message
        
        # Check if error is due to file being locked/in use
        if ($errorMsg -match 'being used by another process' -or $errorMsg -match 'cannot access' -or $errorMsg -match 'File is locked') {
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "ERROR: File is locked or in use" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            Write-Host ""
            Write-Host "The following file is currently being used by another program:" -ForegroundColor Yellow
            Write-Host "  $($v.FullName)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Please close any programs that may be using this file (e.g., VLC, media players, file explorers) and try again." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Script execution stopped." -ForegroundColor Red
            Write-Host ""
            exit 1
        } else {
            Write-Warning ("[SubtitleManager] Failed on '{0}': {1}" -f $v.FullName, $errorMsg)
        }
    }
}

Write-Host ("[SubtitleManager] Done. Updated files: {0}" -f $changed)
