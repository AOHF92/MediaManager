<# =====================================================================
 Convert-And-Swap-To-HEVC10-AAC.ps1
 - Scans MediaRoot for video files
 - Converts anything that is NOT: HEVC Main10 + (p010le OR yuv420p10le) + ALL-AAC audio
 - Moves original to BackupRoot, puts converted .mkv in its place
 - Encoder fallback: NVENC -> AMF -> libx265
 - GUI-ready entrypoint: Invoke-ConvertToHevc10Aac
===================================================================== #>

# Console encodings (UTF-8 for JP paths)
try { chcp 65001 > $null } catch {}
[Console]::InputEncoding  = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding']    = 'utf8'
$PSDefaultParameterValues['Add-Content:Encoding'] = 'utf8'
$PSDefaultParameterValues['Set-Content:Encoding'] = 'utf8'

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ffprobe = 'ffprobe'
$ffmpeg  = 'ffmpeg'

# ----------------- Helpers -----------------

function Get-ToolPath {
    param([Parameter(Mandatory)][string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $cmd) { throw "Required tool '$Name' not found on PATH." }
    return $cmd.Path
}

function New-Directory {
    param([Parameter(Mandatory)][string]$p)
    if (-not (Test-Path -LiteralPath $p)) {
        [void](New-Item -ItemType Directory -Path $p -Force)
    }
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$FullPath
    )
    # Works on Windows PowerShell 5.1 and PowerShell 7+
    $base = (Resolve-Path -LiteralPath $BasePath).Path.TrimEnd('\','/')
    $full = (Resolve-Path -LiteralPath $FullPath).Path

    $baseUri = [Uri]("$base" + [IO.Path]::DirectorySeparatorChar)
    $fullUri = [Uri]$full
    $relUri  = $baseUri.MakeRelativeUri($fullUri)
    $rel     = [Uri]::UnescapeDataString($relUri.ToString()).Replace('/', '\')
    return $rel
}

function Invoke-FFMpeg {
    param([Parameter(Mandatory)][string[]]$ArgsArray)

    # Use call operator with array args (avoids Start-Process quoting edge cases)
    # Redirect stderr to suppress encoder initialization warnings during fallback attempts
    & $ffmpeg @ArgsArray 2>$null
    return $LASTEXITCODE
}

function Test-Convert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Video: codec/profile/pix_fmt
    $vj = & $ffprobe -v error `
        -select_streams v:0 `
        -show_entries stream=codec_name,profile,pix_fmt `
        -of json -- $Path 2>$null

    if (-not $vj) { return $true }
    try { $vjObj = $vj | ConvertFrom-Json } catch { return $true }
    if (-not $vjObj.streams -or $vjObj.streams.Count -lt 1) { return $true }

    $v = $vjObj.streams[0]
    $videoCodec = [string]$v.codec_name
    $videoProfile = [string]$v.profile
    $pixFmt     = [string]$v.pix_fmt

    # Audio: ALL audio streams must be AAC (files with no audio streams are non-compliant)
    $aj = & $ffprobe -v error `
        -select_streams a `
        -show_entries stream=codec_name `
        -of json -- $Path 2>$null

    $aStreams = @()
    if ($aj) {
        try {
            $ajObj = $aj | ConvertFrom-Json
            if ($ajObj.streams) { $aStreams = @($ajObj.streams) }
        } catch {
            # If we can't parse audio info, err on side of converting
            return $true
        }
    }

    $isHevc   = ($videoCodec -eq 'hevc')
    # Match audit script: accept "Main 10" or "Main10" (no space)
    $isMain10 = ($videoProfile -eq 'Main 10' -or $videoProfile -eq 'Main10')

    # Accept NVENC's common 10-bit surface (p010le) and the CPU/x265 surface (yuv420p10le)
    $is10bit420 = ($pixFmt -in @('p010le','yuv420p10le'))

    # Files with no audio streams are non-compliant (must have at least one AAC stream)
    $allAudioAac = $false
    if ($aStreams.Count -gt 0) {
        $nonAac = @($aStreams | Where-Object { $_.codec_name -ne 'aac' })
        if ($nonAac.Count -eq 0) { $allAudioAac = $true }
    }

    $isTarget = $isHevc -and $isMain10 -and $is10bit420 -and $allAudioAac

    # needs conversion = NOT already target
    return -not $isTarget
}

function Get-EncoderPresets {
    param(
        [Parameter(Mandatory)][ValidateSet('hevc_nvenc','hevc_amf','libx265')]
        [string]$Encoder
    )

    switch ($Encoder) {
        'hevc_nvenc' {
            return @(
                "-c:v","hevc_nvenc",
                "-pix_fmt","p010le",
                "-profile:v","main10",
                "-rc","vbr",
                "-cq","19",
                "-b:v","0",
                "-preset","p4",
                "-tune","hq"
            )
        }
        'hevc_amf' {
            # AMF options vary by build/driver; keep it conservative.
            # If it fails at runtime, we’ll fall back to libx265 automatically.
            return @(
                "-c:v","hevc_amf",
                "-profile:v","main10",
                "-pix_fmt","p010le"
            )
        }
        'libx265' {
            return @(
                "-c:v","libx265",
                "-pix_fmt","yuv420p10le",
                "-crf","21",
                "-preset","medium",
                "-x265-params","profile=main10"
            )
        }
    }
}

function Invoke-EncodeWithFallback {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][string]$TempOutPath
    )

    $encoders = @('hevc_nvenc','hevc_amf','libx265')

    foreach ($enc in $encoders) {

        $baseArgs = @(
            "-y",
            "-hide_banner",
            "-loglevel","quiet",
            "-stats",
            "-i", $InputPath,

            # Maps: keep main video + ALL audio + ALL subs + attachments
            "-map","0:v:0",
            "-map","0:a?",
            "-map","0:s?",
            "-map","0:t?",
            "-ignore_unknown",
            "-thread_queue_size","1024",
            "-max_muxing_queue_size","4096"
        )

        $videoArgs = Get-EncoderPresets -Encoder $enc

        $audioArgs = @(
            "-c:a","aac",
            "-ac","2",
            "-ar","48000",
            "-b:a","192k"
        )

        $subsArgs = @(
            "-c:s","copy"
        )

        $ffmpegArgs = @($baseArgs + $videoArgs + $audioArgs + $subsArgs + @($TempOutPath))

        $code = Invoke-FFMpeg -ArgsArray $ffmpegArgs
        if ($code -eq 0 -and (Test-Path -LiteralPath $TempOutPath)) {
            return @{
                Success = $true
                Encoder = $enc
                ExitCode = 0
            }
        }

        # cleanup temp between attempts
        if (Test-Path -LiteralPath $TempOutPath) {
            Remove-Item -LiteralPath $TempOutPath -Force -ErrorAction SilentlyContinue
        }
    }

    return @{
        Success = $false
        Encoder = $null
        ExitCode = 1
    }
}

# ----------------- CORE FUNCTION (GUI entry) -----------------

function Invoke-ConvertToHevc10Aac {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MediaRoot,

        [Parameter(Mandatory)]
        [string]$BackupRoot,

        [switch]$Recurse,
        [switch]$DryRun,

        [string]$FailLog,
        [string]$SummaryCsv
    )

    # Default Recurse to $true unless explicitly provided
    if (-not $PSBoundParameters.ContainsKey('Recurse')) {
        $Recurse = $true
    }

    # Set FailLog relative to script root or current directory if not provided
    if (-not $FailLog) {
        $FailLog = if ($PSScriptRoot) {
            [IO.Path]::Combine($PSScriptRoot, "FailedList.txt")
        } else {
            [IO.Path]::Combine((Get-Location).Path, "FailedList.txt")
        }
    }

    # Set SummaryCsv relative to logs directory if not provided
    $logsRoot = if ($PSScriptRoot) { 
        Join-Path $PSScriptRoot "logs" 
    } else { 
        Join-Path (Get-Location).Path "logs" 
    }
    New-Directory -p $logsRoot

    if (-not $SummaryCsv) {
        $ts = Get-Date -Format "yyyyMMdd_HHmmss"
        $SummaryCsv = Join-Path $logsRoot "ConvertSummary_$ts.csv"
    }
    $summaryRows = New-Object System.Collections.Generic.List[object]

    # Tool existence check early (clear errors)
    [void](Get-ToolPath -Name "ffprobe")
    [void](Get-ToolPath -Name "ffmpeg")

    Write-Host "Scan: $MediaRoot" -ForegroundColor Cyan
    if ($DryRun) {
        Write-Host "Mode: DRY RUN (no files will be converted or moved)" -ForegroundColor Yellow
    }

    if (-not (Test-Path -LiteralPath $MediaRoot -PathType Container)) {
        throw "MediaRoot not found: $MediaRoot"
    }

    New-Directory -p $BackupRoot

    $searchParams = @{
        Path = $MediaRoot
        File = $true
    }
    if ($Recurse) { $searchParams['Recurse'] = $true }

    $videoExts = '.mkv','.mp4','.m4v','.mov','.avi','.webm'
    $videos = @(Get-ChildItem @searchParams |
        Where-Object { $_.Extension.ToLower() -in $videoExts } |
        Where-Object { $_.Name -notmatch '__converting__' })

    if ($videos.Count -eq 0) {
        Write-Host "No video files found under: $MediaRoot" -ForegroundColor Yellow
        return
    }

    foreach ($v in $videos) {

        if (-not (Test-Path -LiteralPath $v.FullName)) {
            Write-Warning "Skipping missing file: $($v.FullName)"
            continue
        }

        if (-not (Test-Convert -Path $v.FullName)) {
            Write-Host "[SKIP] $($v.Name) (already compliant)" -ForegroundColor Green
            continue
        }

        $outDir  = $v.DirectoryName
        $outFile = ([IO.Path]::GetFileNameWithoutExtension($v.Name)) + ".mkv"
        $tmpOut  = Join-Path $outDir ($outFile + ".__converting__.tmp.mkv")
        $final   = Join-Path $outDir $outFile

        $rel    = Get-RelativePathSafe -BasePath $MediaRoot -FullPath $v.FullName
        $bakDst = Join-Path $BackupRoot $rel

        if ($DryRun) {
            Write-Host "[DRY] Would convert: $($v.FullName)" -ForegroundColor Cyan
            Write-Host "      -> output: $final"
            Write-Host "      -> move original to backup: $bakDst"
            $summaryRows.Add([pscustomobject]@{
                Status = "DryRun"
                Path   = $v.FullName
                Title  = [IO.Path]::GetFileNameWithoutExtension($v.Name)
                Reason = "Dry run - no conversion performed"
            })
            continue
        }

        $result = Invoke-EncodeWithFallback -InputPath $v.FullName -TempOutPath $tmpOut

        if ($result.Success) {
            New-Directory -p ([IO.Path]::GetDirectoryName($bakDst))

            Move-Item -LiteralPath $v.FullName -Destination $bakDst -Force
            Move-Item -LiteralPath $tmpOut -Destination $final -Force

            Write-Host "[OK] $($v.Name) (encoder: $($result.Encoder))" -ForegroundColor Green
            $summaryRows.Add([pscustomobject]@{
                Status = "Success"
                Path   = $final
                Title  = [IO.Path]::GetFileNameWithoutExtension($outFile)
                Reason = "Converted successfully using $($result.Encoder)"
            })
        }
        else {
            Write-Host "[FAIL] $($v.Name) (all encoders failed)" -ForegroundColor Red

            if (Test-Path -LiteralPath $tmpOut) {
                Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue
            }

            try {
                $failDir = [IO.Path]::GetDirectoryName($FailLog)
                if ($failDir) { New-Directory -p $failDir }
                Add-Content -LiteralPath $FailLog -Value $v.FullName -ErrorAction SilentlyContinue
            } catch {
                # ignore logging errors
            }

            $summaryRows.Add([pscustomobject]@{
                Status = "Fail"
                Path   = $v.FullName
                Title  = [IO.Path]::GetFileNameWithoutExtension($v.Name)
                Reason = "All encoders failed (nvenc/amf/x265). Check ffmpeg console output for details."
            })
        }
    }

    Write-Host "Done." -ForegroundColor Cyan

    if ($summaryRows.Count -gt 0) {
        $summaryRows | Export-Csv -LiteralPath $SummaryCsv -NoTypeInformation -Encoding UTF8
        Write-Host "Summary CSV saved to: $SummaryCsv" -ForegroundColor Cyan
    }
}

# ----------------- Script entrypoint (CLI behavior) -----------------

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-ConvertToHevc10Aac @PSBoundParameters
}
