[CmdletBinding()]
param(
  [int]$Season,
  [string]$Directory,
  [string[]]$Ext = @("mkv","mp4","srt","ass","ssa","sub","sup"),
  [int]$EpisodeOffset = 0,
  [string]$DefaultSubLang = "",
  [string]$Mal,
  [switch]$WhatIf
)

function Invoke-MalRename {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory = $true)][int]$Season,
      [string]$Directory,
      [string[]]$Ext = @("mkv","mp4","srt","ass","ssa","sub","sup"),
      [int]$EpisodeOffset = 0,
      [string]$DefaultSubLang = "",
      [string]$Mal,
      [switch]$WhatIf
    )

    # Save the directory where the function was launched
    $OriginalDir = $PWD.Path

    try {
        # ------------------------------------------------------
        # 1) Validate directory if provided, otherwise use current directory
        # ------------------------------------------------------
        if (-not $Directory -or $Directory.Trim() -eq "") {
            $Directory = (Get-Location).Path
        }

        $Directory = Resolve-Path -Path $Directory -ErrorAction SilentlyContinue
        if (-not $Directory) {
            Write-Host "ERROR: The directory does not exist: $Directory" -ForegroundColor Red
            throw "Directory not found: $Directory"
        }

        Write-Host "`nUsing directory: $Directory"
        Set-Location $Directory

        # ------------------------------------------------------
        # Summary logging (Old Title, New Title)
        # ------------------------------------------------------
        $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
        if (-not $scriptRoot -or $scriptRoot.Trim() -eq "") {
            $scriptRoot = (Get-Location).Path
        }
        $logsDir = Join-Path $scriptRoot "logs"
        if (-not (Test-Path -LiteralPath $logsDir)) {
            New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
        }

        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $summaryPath = Join-Path $logsDir ("MALRenameSummary_{0}.csv" -f $stamp)

        # Collect rows in memory; we will export at the end
        $summaryRows = New-Object System.Collections.Generic.List[object]

        function Add-SummaryRow {
            param(
                [Parameter(Mandatory=$true)][string]$OldTitle,
                [Parameter(Mandatory=$true)][string]$NewTitle
            )
            $summaryRows.Add([pscustomobject]@{
                'Old Title' = $OldTitle
                'New Title' = $NewTitle
            }) | Out-Null
        }

        # ------------------------------------------------------
        # 2) Validate MAL URL / ID if provided
        # ------------------------------------------------------
        if (-not $Mal -or $Mal.Trim() -eq "") {
            Write-Host "ERROR: MAL ID or URL is required. Provide -Mal parameter." -ForegroundColor Red
            throw "MAL ID or URL is required"
        }

        # Extract MAL numeric ID
        $Mal = $Mal.Trim()
        $MalId = $null

        if ($Mal -match '^\d+$') {
            # Just a number
            $MalId = [int]$Mal
        } else {
            # Try to pull from URL like .../anime/6746/...
            $m = [regex]::Match($Mal, 'anime/(\d+)')
            if ($m.Success) {
                $MalId = [int]$m.Groups[1].Value
            }
        }

        if (-not $MalId) {
            Write-Host "ERROR: Could not extract MAL ID from input: $Mal" -ForegroundColor Red
            throw "Invalid MAL ID or URL: $Mal"
        }

        Write-Host ""
        Write-Host "Using MAL ID: $MalId"

        # ------------------------------------------------------
        # 3) Fetch episode list from Jikan API (like the Python script)
        # ------------------------------------------------------

        function Get-MalEpisodes {
            param(
                [int]$Id
            )

            $page     = 1
            $allEps   = @()

            while ($true) {
                $url = "https://api.jikan.moe/v4/anime/$Id/episodes?page=$page"
                Write-Host "Fetching: $url"

                try {
                    $resp = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 20
                }
                catch {
                    Write-Host "ERROR: Failed to fetch from Jikan: $($_.Exception.Message)" -ForegroundColor Red
                    break
                }

                if (-not $resp) { break }

                if ($resp.data) {
                    $allEps += $resp.data
                }

                $hasNext = $false
                if ($resp.pagination -and $resp.pagination.has_next_page) {
                    $hasNext = [bool]$resp.pagination.has_next_page
                }

                if (-not $hasNext) {
                    break
                }

                $page++
                Start-Sleep -Milliseconds 500  # be polite with the API
            }

            return $allEps
        }

        Write-Host ""
        Write-Host "Downloading episode titles from Jikan..."
        $episodes = Get-MalEpisodes -Id $MalId

        if (-not $episodes -or $episodes.Count -eq 0) {
            Write-Host "ERROR: No episodes returned from Jikan. Check MAL ID or try again later." -ForegroundColor Red
            return
        }

        Write-Host ("Fetched {0} episodes from Jikan." -f $episodes.Count)

        # ---------- helpers ----------
         function Remove-InvalidChars {
             param(
                [string]$s
             )
          if ($null -eq $s) { return "" }
          $bad = [IO.Path]::GetInvalidFileNameChars() -join ''
          $re = "[{0}]" -f [Regex]::Escape($bad)
          ($s -replace $re,'-').Trim()
        }
         function Format-StringLength {
             param(
                [string]$s,
                [int]$max
             )
          if ($null -eq $s) { return "" }
          if ($s.Length -le $max) { return $s }
          if ($max -le 1) { return "…" }
          return ($s.Substring(0, $max - 1) + "…")
        }

        # ------------------------------------------------------
        # 4) Build episode -> titles map (from Jikan response)
        # ------------------------------------------------------
        $map = @{}

        foreach ($e in $episodes) {
            # Python version used mal_id or episode, whichever was present
            $epNum  = $null
            if ($e.mal_id) {
                $epNum = [int]$e.mal_id
            } elseif ($e.episode) {
                $epNum = [int]$e.episode
            }

            if ($null -eq $epNum) { continue }

            $titleEn = $e.title
            $titleJp = $e.title_japanese

            $en = Remove-InvalidChars $titleEn
            $jp = Remove-InvalidChars $titleJp

            $map[$epNum] = @{ en=$en; jp=$jp }
        }

        if ($map.Count -eq 0) {
            Write-Host "ERROR: Episode map is empty; cannot continue." -ForegroundColor Red
            return
        }

        # ------------------------------------------------------
        # 5) Episode extractor & language detection (your original logic)
        # ------------------------------------------------------
         function Get-EpisodeNumber {
             param(
                [string]$baseName
             )
          # 0) Ep / Episode (ignore decimal like 12.5 -> 12)
          $m = [regex]::Match($baseName, '(?i)(?:^|[ \.\-_])(?:ep|episode)[ \._-]*0*(?<e>\d{1,3})(?:\.\d+)?(?=$|[ \.\-_])')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 1) SxxEyy
          $m = [regex]::Match($baseName, 'S\d{1,2}E(?<e>\d{2})', 'IgnoreCase')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 2) [02]
          $m = [regex]::Match($baseName, '(?<=\[)(?<e>\d{2})(?=\])')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 3) 02 right before '('
          $m = [regex]::Match($baseName, '(?<e>\d{2})(?=[\s._\-\[\(]*\()')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 4) two digits between common separators
          $m = [regex]::Match($baseName, '(?<=[\s._\-\[\]])(?<e>\d{2})(?=[\s._\-\[\]]|$)')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 5) leading number
          $m = [regex]::Match($baseName, '^(?<e>\d{1,3})\b')
          if ($m.Success) { return [int]$m.Groups['e'].Value }

          # 6) fallback: choose the two-digit number closest to '(' with safe borders
          $paren = $baseName.IndexOf('('); if ($paren -lt 0) { $paren = $baseName.Length }
          $cands = @()
          foreach ($mm in [regex]::Matches($baseName, '\d{2}')) {
            $i = $mm.Index; $j = $i + 1
            $prev = if ($i -gt 0) { $baseName[$i-1] } else { [char]0 }
            $next = if ($j + 1 -le $baseName.Length) { $baseName[$j] } else { [char]0 }
            $okPrev = ($i -eq 0) -or ($prev -match '[\s._\-\[\]]')
            $okNext = ($j -eq $baseName.Length-1) -or ($next -match '[\s._\-\[\]]') -or ($next -eq '(')
            if ($okPrev -and $okNext) {
              $dist = [math]::Abs($i - $paren)
              $cands += [pscustomobject]@{ ep=[int]$mm.Value; dist=$dist; idx=$i }
            }
          }
          if ($cands.Count) { return ($cands | Sort-Object dist, idx)[0].ep }
          return $null
        }

         function Get-Language {
             param(
                [string]$name
             )
          $m = [regex]::Match($name, '(?:^|[ \.\-_])(?<lang>en|eng|ja|jpn|jp|sc|chs|cn|tc|cht|kr|ko|es|spa|pt|por|fr|fre|de|ger|it|ru|rus|pl|vi|id|ind)(?=[ \.\-_]|$)', 'IgnoreCase')
          if (-not $m.Success) { return $null }
          switch ($m.Groups['lang'].Value.ToLower()){
            'en' {'EN'} 'eng' {'EN'}
            'ja' {'JA'} 'jp' {'JA'} 'jpn' {'JA'}
            'sc' {'SC'} 'chs' {'SC'} 'cn' {'SC'}
            'tc' {'TC'} 'cht' {'TC'}
            'kr' {'KO'} 'ko' {'KO'}
            'es' {'ES'} 'spa' {'ES'}
            'pt' {'PT'} 'por' {'PT'}
            'fr' {'FR'} 'fre' {'FR'}
            'de' {'DE'} 'ger' {'DE'}
            'it' {'IT'}
            'ru' {'RU'} 'rus' {'RU'}
            'pl' {'PL'}
            'vi' {'VI'}
            'id' {'ID'} 'ind' {'ID'}
            default { $null }
          }
        }

        $subtitleExts = @(".srt",".ass",".ssa",".sub",".sup")

        # ------------------------------------------------------
        # 6) Scan files & build metadata
        # ------------------------------------------------------
        $files = Get-ChildItem -File |
          Where-Object { $Ext -contains $_.Extension.TrimStart('.') } |
          ForEach-Object {
            $rawEp = Get-EpisodeNumber $_.BaseName
            [pscustomobject]@{
              File    = $_
              RawEp   = $rawEp
              Ep      = if ($null -ne $rawEp) { [int]($rawEp - $EpisodeOffset) } else { $null }
              Ext     = $_.Extension.ToLower()
              Lang    = if ($subtitleExts -contains $_.Extension.ToLower()) { Get-Language $_.BaseName } else { $null }
            }
        }

        $groups = $files | Group-Object { "{0}|{1}" -f $_.Ep, $_.Ext }

        # ------------------------------------------------------
        # 7) Rename loop (same logic as before)
        # ------------------------------------------------------
        foreach ($g in $groups) {
          $items = $g.Group
          foreach ($entry in $items) {
            $f   = $entry.File
            $ext = $entry.Ext
            $dir = $f.DirectoryName
            $ep  = $entry.Ep

            if ($null -eq $ep) {
              Write-Warning "Could not detect episode number in: $($f.Name)"
              continue
            }
            if (-not $map.ContainsKey($ep)) {
              Write-Warning "Mapped episode $ep not found in Jikan data (raw was $($entry.RawEp)). Adjust -EpisodeOffset? Skipping: $($f.Name)"
              continue
            }

            $en = $map[$ep].en
            $jp = $map[$ep].jp

             function New-EpisodeName {
                 param(
                    $en2,
                    $jp2
                 )
              "{0}. {1} ({2}) - S{3:00}E{4:00}" -f $ep,$en2,$jp2,$Season,$ep
            }

            $base = New-EpisodeName $en $jp

            # smart subtitle tag: only if multiple variants for this (ep, ext)
            $langSuffix = ""
            if ($subtitleExts -contains $ext) {
              if ($items.Count -gt 1) {
                $langSuffix = $entry.Lang
                if ([string]::IsNullOrWhiteSpace($langSuffix) -and -not [string]::IsNullOrWhiteSpace($DefaultSubLang)) {
                  $langSuffix = $DefaultSubLang
                }
              } else {
                $langSuffix = ""
              }
            }

            $newName = "$base$langSuffix$ext"

            # long-path guard
            $maxTotal   = 259
            $fullTarget = Join-Path $dir $newName
            if ($fullTarget.Length -gt $maxTotal) {
              $allowBase = $maxTotal - ($dir.Length + 1 + $ext.Length)
              if ($allowBase -lt 20) { $allowBase = 20 }

              $trimmed = $false
              $enCap = [Math]::Min($en.Length, 120)
              $jpCap = [Math]::Min($jp.Length, 80)

              $enTry  = Format-StringLength $en $enCap
              $baseTry = New-EpisodeName $enTry $jp
              if ( ($baseTry + $langSuffix).Length -le $allowBase ) {
                $base = $baseTry; $trimmed = $true
              } else {
                $jpTry  = Format-StringLength $jp $jpCap
                $baseTry2 = New-EpisodeName $enTry $jpTry
                if ( ($baseTry2 + $langSuffix).Length -le $allowBase ) {
                  $base = $baseTry2; $trimmed = $true
                } else {
                  $enAllow = [int]([Math]::Floor($allowBase * 0.60))
                  $jpAllow = $allowBase - $enAllow - 5
                  if ($jpAllow -lt 10) { $jpAllow = 10 }
                  if ($enAllow -lt 10) { $enAllow = 10 }
                  $enTry2 = Format-StringLength $en $enAllow
                  $jpTry2 = Format-StringLength $jp $jpAllow
                  $base = New-EpisodeName $enTry2 $jpTry2
                  if ( ($base + $langSuffix).Length -gt $allowBase ) {
                    $base = Format-StringLength ($base + $langSuffix) $allowBase
                    $langSuffix = ""
                  }
                  $trimmed = $true
                }
              }
              $newName = "$base$langSuffix$ext"
              if ($trimmed) { Write-Host "[Trimmed] $($f.Name) -> $newName" }
            }

            if ($WhatIf) {
              Write-Host "[DRY] $($f.Name)  ->  $newName"
              Add-SummaryRow -OldTitle $f.Name -NewTitle $newName
            } else {
              if (Test-Path -LiteralPath (Join-Path $dir $newName)) {
                Write-Warning "Destination exists, skipping: $newName"
              } else {
                Rename-Item -LiteralPath $f.FullName -NewName $newName
                Add-SummaryRow -OldTitle $f.Name -NewTitle $newName
              }
            }
          }
        }
    }
    finally {
        # ------------------------------------------------------
        # Write summary CSV (per run)
        # ------------------------------------------------------
        try {
            if ($summaryRows.Count -gt 0) {
                # Use UTF-8 with BOM for Excel compatibility
                $utf8WithBom = New-Object System.Text.UTF8Encoding $true
                $csvContent = $summaryRows | ConvertTo-Csv -NoTypeInformation
                [System.IO.File]::WriteAllLines($summaryPath, $csvContent, $utf8WithBom)
                Write-Host ""
                Write-Host "Summary CSV: $summaryPath" -ForegroundColor Green
            } else {
                Write-Host ""
                Write-Host "No renames performed; summary not written." -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Failed to write summary CSV: $($_.Exception.Message)"
        }

        # Restore original directory no matter what happens
        Set-Location $OriginalDir
        Write-Host ""
        Write-Host "Returned to original directory: $OriginalDir" -ForegroundColor Cyan
    }
}

# --- Script entrypoint: only runs when the script is executed, not dot-sourced ---

if ($MyInvocation.InvocationName -ne '.') {
    # Script entrypoint - if Season not provided, prompt (interactive mode)
    if (-not $PSBoundParameters.ContainsKey('Season')) {
        try {
            $Season = [int](Read-Host "Enter Season number (1,2,3,...)")
        } catch {
            Write-Host "ERROR: Season number is required." -ForegroundColor Red
            exit 1
        }
    }
    
    # If Directory or Mal not provided, prompt (interactive mode)
    if (-not $PSBoundParameters.ContainsKey('Directory')) {
        Write-Host ""
        Write-Host "Please enter the full directory path that contains your anime files:"
        $Directory = Read-Host "> "
    }
    
    if (-not $PSBoundParameters.ContainsKey('Mal')) {
        Write-Host ""
        Write-Host "Enter MAL anime URL or numeric ID (example: https://myanimelist.net/anime/6746/ or 6746):"
        $Mal = Read-Host "> "
    }

    Invoke-MalRename `
        -Season         $Season `
        -Directory      $Directory `
        -Ext            $Ext `
        -EpisodeOffset  $EpisodeOffset `
        -DefaultSubLang $DefaultSubLang `
        -Mal            $Mal `
        -WhatIf:$WhatIf
}
