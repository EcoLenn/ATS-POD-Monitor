# Compares carrier POD (Proof of Delivery) shipment data against SYSPRO ArInvoice records.
#
# Usage examples:
# .\Compare-PodData.ps1
# .\Compare-PodData.ps1 -CsvPath "P:\Documents\ecotrend_shipments_starting_20260101_utc.csv"
# .\Compare-PodData.ps1 -StartDate "2026-01-01" -OutputFolder "C:\Automation\ATS POD Monitor\Output"
#
# What it does:
# 1. Reads the carrier shipment CSV (account_province, tracking_number, created_at_utc, ref).
# 2. Queries SYSPRO ArInvoice for Customer, Invoice, InvoiceDate, ProofOfDelivery, PodEntryDate, PodReference
#    where InvoiceDate >= StartDate.
# 3. Matches CSV `ref` to ArInvoice `Invoice` (Invoice is stored zero-padded to 15 chars; matched on the
#    trimmed/zero-stripped value).
# 4. Converts each shipment's created_at_utc (UTC) to the account's local time using account_province
#    (ON = Eastern Time, BC = Pacific Time). PodEntryDate in ArInvoice is already stored in that same
#    account-local time, so no further conversion is applied to it.
# 5. Writes three reports to OutputFolder:
#      - PodComparison_<timestamp>.csv   Matched rows with a recorded PodReference: local POD time vs.
#                                         PodEntryDate, and the difference in minutes/hours.
#      - PodUpdate_<timestamp>.sql       UPDATE statements for matched invoices that have no PodReference
#                                         recorded yet (PodReference = ' '), following the existing manual
#                                         update pattern used for ArInvoice.
#      - PodUnmatched_<timestamp>.csv    CSV rows whose `ref` did not match any Invoice returned by the
#                                         query (InvoiceDate >= StartDate) - needs manual review, no SQL
#                                         is generated since the WHERE clause would match zero rows.

param(
    [string]$CsvPath = "P:\Documents\ecotrend_shipments_starting_20260101_utc.csv",
    [string]$Server = "EE01-VMSRV02",
    [string]$Database = "SYSPROCompanyE",
    [string]$StartDate = "2026-01-01",
    [string]$OutputFolder = "C:\Automation\ATS POD Monitor\Output"
)

$ErrorActionPreference = "Stop"

# Windows time zone IDs (handle DST automatically). Confirmed with the business:
# ON -> Eastern Time, BC -> Pacific Time. PodEntryDate in ArInvoice is already recorded
# in this same account-local zone, so it is compared as-is (no conversion).
$AccountTimeZones = @{
    "ON" = "Eastern Standard Time"
    "BC" = "Pacific Standard Time"
}

function Get-NormalizedInvoiceKey {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    return $Value.Trim().TrimStart("0")
}

if (-not (Test-Path $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$comparisonFile = Join-Path $OutputFolder "PodComparison_$timestamp.csv"
$updateSqlFile = Join-Path $OutputFolder "PodUpdate_$timestamp.sql"
$unmatchedFile = Join-Path $OutputFolder "PodUnmatched_$timestamp.csv"

# --- Load carrier CSV -------------------------------------------------------
# Only account_province (for time zone), tracking_number, created_at_utc, and ref are used.
# account_name, status, and po are not needed for this comparison.
Write-Host "Reading CSV: $CsvPath"
$csvRows = Import-Csv -Path $CsvPath

Write-Host "CSV rows loaded: $($csvRows.Count)"

# --- Query SYSPRO ArInvoice --------------------------------------------------
Write-Host "Querying $Database on $Server (InvoiceDate >= $StartDate)..."

$query = @"
SELECT Customer, Invoice, InvoiceDate, ProofOfDelivery, PodEntryDate, PodReference
FROM dbo.ArInvoice
WHERE InvoiceDate >= @StartDate
"@

$connectionString = "Server=$Server;Database=$Database;Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

$sqlByKey = @{}

try {
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.CommandTimeout = 0
    $null = $command.Parameters.AddWithValue("@StartDate", [datetime]$StartDate)

    $reader = $command.ExecuteReader()
    $table = New-Object System.Data.DataTable
    $table.Load($reader)

    Write-Host "ArInvoice rows returned: $($table.Rows.Count)"

    foreach ($row in $table.Rows) {
        $invoiceKey = Get-NormalizedInvoiceKey ([string]$row["Invoice"])
        if ($invoiceKey -eq "") { continue }

        $podEntryDate = $null
        if ($row["PodEntryDate"] -isnot [DBNull]) { $podEntryDate = [datetime]$row["PodEntryDate"] }

        $sqlRecord = [PSCustomObject]@{
            Customer        = [string]$row["Customer"]
            Invoice         = ([string]$row["Invoice"]).Trim()
            InvoiceDate     = $row["InvoiceDate"]
            ProofOfDelivery = [string]$row["ProofOfDelivery"]
            PodEntryDate    = $podEntryDate
            PodReference    = ([string]$row["PodReference"]).Trim()
        }

        if (-not $sqlByKey.ContainsKey($invoiceKey)) {
            $sqlByKey[$invoiceKey] = New-Object System.Collections.Generic.List[object]
        }
        $sqlByKey[$invoiceKey].Add($sqlRecord)
    }
}
finally {
    if ($connection.State -eq "Open") { $connection.Close() }
}

# --- Compare -----------------------------------------------------------------
$comparisonResults = New-Object System.Collections.Generic.List[object]
$updateStatements = New-Object System.Collections.Generic.List[string]
$unmatchedResults = New-Object System.Collections.Generic.List[object]

$matchedWithPod = 0
$matchedNeedingUpdate = 0
$unmatchedCount = 0

foreach ($csvRow in $csvRows) {
    $province = $csvRow.account_province
    $ref = $csvRow.ref
    $trackingNumber = $csvRow.tracking_number

    if (-not $AccountTimeZones.ContainsKey($province)) {
        Write-Warning "No time zone mapping for account_province '$province' (ref=$ref, tracking=$trackingNumber) - skipping."
        continue
    }

    $utcDt = [datetime]::Parse($csvRow.created_at_utc, [System.Globalization.CultureInfo]::InvariantCulture, `
        [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)

    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($AccountTimeZones[$province])
    $localPodTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcDt, $tz)

    $invoiceKey = Get-NormalizedInvoiceKey $ref

    if ($sqlByKey.ContainsKey($invoiceKey)) {
        foreach ($sqlRecord in $sqlByKey[$invoiceKey]) {
            $hasPodReference = -not [string]::IsNullOrWhiteSpace($sqlRecord.PodReference)

            if ($hasPodReference -and $sqlRecord.PodEntryDate) {
                $diff = $sqlRecord.PodEntryDate - $localPodTime
                $podEntryDateHasTime = ($sqlRecord.PodEntryDate.TimeOfDay -ne [timespan]::Zero)
                $reviewFlag = ""
                if (-not $podEntryDateHasTime) { $reviewFlag = "PodEntryDate has no time-of-day (00:00:00) - diff reflects date only, not true delay" }
                if ([math]::Abs($diff.TotalDays) -gt 30) { $reviewFlag = "Diff exceeds 30 days - check for a data entry error in PodEntryDate (e.g. wrong year)" }

                $comparisonResults.Add([PSCustomObject]@{
                    Ref                 = $ref
                    Invoice             = $sqlRecord.Invoice
                    Customer            = $sqlRecord.Customer
                    AccountProvince     = $province
                    TrackingNumber      = $trackingNumber
                    CreatedAtUtc        = $csvRow.created_at_utc
                    PodTimeLocal        = $localPodTime.ToString("yyyy-MM-dd HH:mm:ss")
                    PodEntryDate        = $sqlRecord.PodEntryDate.ToString("yyyy-MM-dd HH:mm:ss")
                    PodEntryDateHasTime = $podEntryDateHasTime
                    DiffMinutes         = [math]::Round($diff.TotalMinutes, 1)
                    DiffHours           = [math]::Round($diff.TotalHours, 2)
                    ProofOfDelivery     = $sqlRecord.ProofOfDelivery
                    PodReferenceSql     = $sqlRecord.PodReference
                    TrackingMatchesPod  = ($sqlRecord.PodReference -eq $trackingNumber)
                    ReviewFlag          = $reviewFlag
                })
                $matchedWithPod++
            }
            else {
                # Matched invoice, but PodReference/ProofOfDelivery not recorded yet - needs an UPDATE.
                $refEscaped = $ref -replace "'", "''"
                $trackingEscaped = $trackingNumber -replace "'", "''"
                $podEntryDateSql = $localPodTime.ToString("M/d/yyyy h:mm:ss tt", [System.Globalization.CultureInfo]::InvariantCulture)

                $statement = @"
UPDATE dbo.ArInvoice
SET PodReference = '$trackingEscaped',
    PodEntryDate = '$podEntryDateSql',
    ProofOfDelivery = 'Y'
WHERE Invoice = REPLACE(SPACE(15 - LEN('$refEscaped')) + '$refEscaped', ' ', '0')
AND PodReference = ' ';
"@
                $updateStatements.Add($statement)
                $matchedNeedingUpdate++
            }
        }
    }
    else {
        $unmatchedResults.Add([PSCustomObject]@{
            Ref             = $ref
            AccountProvince = $province
            TrackingNumber  = $trackingNumber
            CreatedAtUtc    = $csvRow.created_at_utc
            PodTimeLocal    = $localPodTime.ToString("yyyy-MM-dd HH:mm:ss")
            Reason          = "No ArInvoice.Invoice match for InvoiceDate >= $StartDate"
        })
        $unmatchedCount++
    }
}

# --- Write outputs -------------------------------------------------------------
$comparisonResults | Export-Csv -Path $comparisonFile -NoTypeInformation -Encoding UTF8

if ($updateStatements.Count -gt 0) {
    $header = "-- Generated by Compare-PodData.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n" +
              "-- Updates ArInvoice.PodReference/PodEntryDate/ProofOfDelivery from carrier POD data`r`n" +
              "-- Source CSV: $CsvPath`r`n" +
              "-- Review before running against $Database on $Server.`r`n`r`n"
    ($header + ($updateStatements -join "`r`n`r`n")) | Out-File -FilePath $updateSqlFile -Encoding UTF8
}

$unmatchedResults | Export-Csv -Path $unmatchedFile -NoTypeInformation -Encoding UTF8

# --- Summary -------------------------------------------------------------------
$noTimeOfDayCount = ($comparisonResults | Where-Object { -not $_.PodEntryDateHasTime }).Count
$flaggedCount = ($comparisonResults | Where-Object { $_.ReviewFlag -ne "" }).Count

Write-Host ""
Write-Host "Matched, POD already recorded : $matchedWithPod  -> $comparisonFile"
Write-Host "  - PodEntryDate has no time-of-day (00:00:00): $noTimeOfDayCount ($([math]::Round(100*$noTimeOfDayCount/[math]::Max($matchedWithPod,1),1))%) - diff is date-only for these, not a true delay"
Write-Host "  - Flagged for review (see ReviewFlag column): $flaggedCount"
Write-Host "Matched, POD needs update     : $matchedNeedingUpdate  -> $updateSqlFile"
Write-Host "No matching invoice found     : $unmatchedCount  -> $unmatchedFile"
