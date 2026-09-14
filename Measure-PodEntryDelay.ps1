# Measures the delay between an invoice's recorded PodEntryDate and the moment the ATS
# integration actually wrote it to SYSPRO ArInvoice, using the ATS WorkingDirectory\Log
# .debug files as the source (not SYSPRO itself).
#
# Each .debug log line that runs the "query to update orders with ATS tracking numbers"
# statement looks like:
#
#   Aug22 22:13:23 - [Int32 ExecuteSQL(System.String, System.String)][:0] /*...*/
#   update ArInvoice
#   set PodReference = '805231804',
#       PodEntryDate = '8/21/2026 12:26:28 PM',
#       ProofOfDelivery = 'Y'
#   where Invoice = REPLACE(SPACE(15 - LEN('424230')) + '424230', ' ','0')
#   and PodReference = ' '
#
# "Aug22 22:13:23" is when the update actually ran (UpdatedDateTime) - the year is taken
# from the log file name (Log_YYYYMMDD.debug), since the line itself has no year.
# "PodEntryDate" is the POD time as reported by the carrier/ATS (PodEntryDate).
# HoursDiff = UpdatedDateTime - PodEntryDate, i.e. how many hours passed between the POD
# event and it landing in SYSPRO.
#
# Usage:
#   .\Measure-PodEntryDelay.ps1
#   .\Measure-PodEntryDelay.ps1 -StartDate "2026-04-01" -EndDate "2026-08-28"
#
# Run monthly - just re-run with no arguments to pick up "April 1 through today" again,
# or pass -StartDate/-EndDate to cover a different window.

param(
    [string[]]$LogFolders = @(
        "\\Ee01-vmsrv02\e$\ATSEcotrend\WorkingDirectory\Log",
        "\\Ee01-vmsrv02\e$\ATSEcotrendOnt\WorkingDirectory\Log"
    ),
    [datetime]$StartDate = "2026-04-01",
    [datetime]$EndDate = (Get-Date),
    [string]$OutputFolder = "C:\Automation\ATS POD Monitor\Output"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$headerPattern = '(?m)^(?<mon>[A-Za-z]{3})(?<day>\d{1,2}) (?<time>\d{2}:\d{2}:\d{2}) - '
$fieldsPattern = 'set\s+PodReference\s*=\s*''(?<podref>[^'']*)'',\s*PodEntryDate\s*=\s*''(?<podentry>[^'']*)'',\s*ProofOfDelivery\s*=\s*''Y''.*?LEN\(''(?<invoice>[^'']*)''\)'
$singleline = [System.Text.RegularExpressions.RegexOptions]::Singleline

$results = New-Object System.Collections.Generic.List[object]
$filesProcessed = 0
$entriesFound = 0
$parseFailures = 0

foreach ($folder in $LogFolders) {
    if (-not (Test-Path $folder)) {
        Write-Warning "Log folder not found, skipping: $folder"
        continue
    }

    $source = Split-Path (Split-Path $folder -Parent) -Parent | Split-Path -Leaf   # e.g. ATSEcotrend / ATSEcotrendOnt

    $logFiles = Get-ChildItem -Path $folder -Filter "Log_*.debug" | Where-Object {
        $_.Name -match '^Log_(\d{4})(\d{2})(\d{2})\.debug$'
    } | ForEach-Object {
        $null = $_.Name -match '^Log_(\d{4})(\d{2})(\d{2})\.debug$'
        [PSCustomObject]@{
            File     = $_
            FileDate = Get-Date -Year $matches[1] -Month $matches[2] -Day $matches[3] -Hour 0 -Minute 0 -Second 0
            Year     = $matches[1]
        }
    } | Where-Object { $_.FileDate.Date -ge $StartDate.Date -and $_.FileDate.Date -le $EndDate.Date }

    Write-Host "$source : $($logFiles.Count) log file(s) in range $($StartDate.ToString('yyyy-MM-dd')) - $($EndDate.ToString('yyyy-MM-dd'))"

    foreach ($lf in $logFiles) {
        $filesProcessed++
        $content = Get-Content -Path $lf.File.FullName -Raw
        if ([string]::IsNullOrEmpty($content)) { continue }

        $headerMatches = [regex]::Matches($content, $headerPattern)
        for ($i = 0; $i -lt $headerMatches.Count; $i++) {
            $entryStart = $headerMatches[$i].Index
            $entryEnd = if ($i + 1 -lt $headerMatches.Count) { $headerMatches[$i + 1].Index } else { $content.Length }
            $entryLength = $entryEnd - $entryStart
            if ($entryLength -le 0) { continue }

            # Cheap pre-check before running the heavier field regex.
            $peek = $content.Substring($entryStart, [Math]::Min($entryLength, 400))
            if ($peek -notmatch 'update\s+ArInvoice') { continue }

            $entryText = $content.Substring($entryStart, $entryLength)
            if ($entryText -notmatch 'PodEntryDate') { continue }

            $fieldsMatch = [regex]::Match($entryText, $fieldsPattern, $singleline)
            if (-not $fieldsMatch.Success) { continue }

            $entriesFound++

            $mon = $headerMatches[$i].Groups['mon'].Value
            $day = $headerMatches[$i].Groups['day'].Value
            $time = $headerMatches[$i].Groups['time'].Value

            try {
                $updatedDateTime = [datetime]::ParseExact(
                    "$mon $day $($lf.Year) $time",
                    "MMM d yyyy HH:mm:ss",
                    [System.Globalization.CultureInfo]::InvariantCulture)

                $podEntryDate = [datetime]::Parse(
                    $fieldsMatch.Groups['podentry'].Value,
                    [System.Globalization.CultureInfo]::InvariantCulture)
            }
            catch {
                $parseFailures++
                continue
            }

            $diffHours = ($updatedDateTime - $podEntryDate).TotalHours

            $reviewFlag = ""
            if ($diffHours -lt 0) { $reviewFlag = "PodEntryDate is after the update time - check source data" }

            $results.Add([PSCustomObject]@{
                Source           = $source
                Month            = $updatedDateTime.ToString("yyyy-MM")
                Date             = $updatedDateTime.ToString("yyyy-MM-dd")
                UpdatedTime      = $updatedDateTime.ToString("HH:mm:ss")
                Invoice          = $fieldsMatch.Groups['invoice'].Value
                PodReference     = $fieldsMatch.Groups['podref'].Value
                PodEntryDate     = $podEntryDate.ToString("yyyy-MM-dd HH:mm:ss")
                UpdatedDateTime  = $updatedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                HoursDiff        = [math]::Round($diffHours, 2)
                IsFirstAttempt   = $false
                ReviewFlag       = $reviewFlag
                LogFile          = $lf.File.Name
            })
        }
    }
}

foreach ($group in ($results | Group-Object Source, Invoice, PodEntryDate)) {
    $ordered = $group.Group | Sort-Object UpdatedDateTime
    for ($j = 0; $j -lt $ordered.Count; $j++) {
        $ordered[$j].IsFirstAttempt = ($j -eq 0)
    }
}

$sorted = $results | Sort-Object Month, Date, UpdatedTime, Invoice

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutputFolder "PodEntryDelay_$timestamp.csv"
$sorted | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

$firstAttemptRows = $sorted | Where-Object IsFirstAttempt
$reattemptCount = $sorted.Count - $firstAttemptRows.Count
$flaggedCount = ($sorted | Where-Object { $_.ReviewFlag -ne "" }).Count

Write-Host ""
Write-Host "Log files processed : $filesProcessed"
Write-Host "POD update entries  : $entriesFound  (first attempts: $($firstAttemptRows.Count), re-attempts on already-set rows: $reattemptCount)"
Write-Host "Flagged for review (ReviewFlag column, negative HoursDiff) : $flaggedCount"
if ($parseFailures -gt 0) { Write-Warning "Entries with unparseable dates (skipped): $parseFailures" }
if ($firstAttemptRows.Count -gt 0) {
    $avg = [math]::Round(($firstAttemptRows | Measure-Object -Property HoursDiff -Average).Average, 2)
    $max = [math]::Round(($firstAttemptRows | Measure-Object -Property HoursDiff -Maximum).Maximum, 2)
    $min = [math]::Round(($firstAttemptRows | Measure-Object -Property HoursDiff -Minimum).Minimum, 2)
    Write-Host "Avg / Min / Max hours delay (first attempts only) : $avg / $min / $max"
}
Write-Host "Report -> $outFile"
