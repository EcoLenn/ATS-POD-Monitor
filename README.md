# ATS POD Monitor

## Purpose
Compares carrier Proof-of-Delivery (POD) shipment data against SYSPRO `ArInvoice` records to
check how our recorded POD time compares to what the shipping company reports, and to generate
the SQL needed to record POD data for invoices that don't have it yet.

## Inputs
| Source | Details |
| --- | --- |
| Carrier CSV | `ecotrend_shipments_starting_20260101_utc.csv` - columns used: `account_province`, `tracking_number`, `created_at_utc`, `ref`. `account_name`, `status`, and `po` are not used. |
| SYSPRO query | `EE01-VMSRV02.SYSPROCompanyE.dbo.ArInvoice` - `Customer, Invoice, InvoiceDate, ProofOfDelivery, PodEntryDate, PodReference` where `InvoiceDate >= StartDate` |

The CSV `ref` column is matched against `ArInvoice.Invoice`. `Invoice` is stored zero-padded to
15 characters (e.g. `000424651/424652`), so matching trims and strips leading zeros from both
sides before comparing.

## Time zone handling
`created_at_utc` in the carrier CSV is in UTC. Per account:

| account_province | Time zone |
| --- | --- |
| ON (Ecotrend Ecologics) | Eastern Time |
| BC (Ecotrend Ecologics Ltd.) | Pacific Time |

`ArInvoice.PodEntryDate` is already recorded in that same account-local time zone, so it is
compared directly against the converted local POD time - no further conversion is applied to it.
Time zone conversion uses Windows time zone IDs (`Eastern Standard Time` / `Pacific Standard
Time`), which apply DST automatically.

## What the script does
`Compare-PodData.ps1`:
1. Reads the carrier CSV.
2. Queries `ArInvoice` for the given `StartDate`.
3. Matches each CSV row to an `ArInvoice.Invoice` by `ref`.
4. For matches that already have a `PodReference` recorded, computes the difference between our
   recorded `PodEntryDate` and the carrier's POD time (converted to account-local), in minutes and
   hours.
5. For matches where `PodReference` is still blank (`' '`), generates an `UPDATE` statement setting
   `PodReference` (= tracking number), `PodEntryDate` (= carrier POD time, converted to
   account-local), and `ProofOfDelivery = 'Y'` - following the existing manual update pattern:
   ```sql
   UPDATE dbo.ArInvoice
   SET PodReference = '805241652',
       PodEntryDate = '8/24/2026 3:02:19 PM',
       ProofOfDelivery = 'Y'
   WHERE Invoice = REPLACE(SPACE(15 - LEN('424651/424652')) + '424651/424652', ' ', '0')
   AND PodReference = ' ';
   ```
6. For CSV rows whose `ref` has no matching `Invoice` at all (for the given `StartDate`), lists
   them separately for manual review - no SQL is generated for these since the `WHERE` clause
   would match zero rows.

## Outputs
Written to `Output\` (git-ignored):
| File | Contents |
| --- | --- |
| `PodComparison_<timestamp>.csv` | Matched rows with a recorded POD: our time vs. the carrier's time, and the difference |
| `PodUpdate_<timestamp>.sql` | `UPDATE` statements for matched invoices missing a `PodReference` - review before running |
| `PodUnmatched_<timestamp>.csv` | CSV rows with no matching `ArInvoice.Invoice` - needs manual review |

`PodComparison` includes two data-quality flags because of what a first run against live data found:
* `PodEntryDateHasTime` - `false` when `PodEntryDate` has no time-of-day (`00:00:00`). As of this
  writing, **~61%** of existing `PodEntryDate` values are date-only, so their `DiffMinutes`/`DiffHours`
  reflects a date-vs-time-of-day gap, not a true delay - filter these out before drawing conclusions
  about POD entry speed.
* `ReviewFlag` - set when a diff exceeds 30 days, which usually means a data entry error in
  `PodEntryDate` (e.g. invoice 417231 has `PodEntryDate = 2260-05-25`, almost certainly a typo for
  `2026-05-25`).

## Usage
```powershell
.\Compare-PodData.ps1
.\Compare-PodData.ps1 -CsvPath "P:\Documents\ecotrend_shipments_starting_20260101_utc.csv"
.\Compare-PodData.ps1 -StartDate "2026-01-01" -OutputFolder "C:\Automation\ATS POD Monitor\Output"
```

## Parameters
| Parameter | Default |
| --- | --- |
| `CsvPath` | `P:\Documents\ecotrend_shipments_starting_20260101_utc.csv` |
| `Server` | `EE01-VMSRV02` |
| `Database` | `SYSPROCompanyE` |
| `StartDate` | `2026-01-01` |
| `OutputFolder` | `C:\Automation\ATS POD Monitor\Output` |

## Dependencies
* PowerShell
* SQL Server access to `EE01-VMSRV02` (Windows Integrated Security)
* Read access to the carrier CSV path

## Notes / assumptions
* Every row that matches an `Invoice` but has more than one `ArInvoice` record for that key
  produces one comparison/update row per record.
* Always review `PodUpdate_*.sql` before running it against SYSPRO.
* If the same `ref` appears more than once in the CSV while `PodReference` is still blank, only
  the first `UPDATE` run will actually change data - the `AND PodReference = ' '` guard makes the
  rest no-ops on a re-run, matching the existing manual process.

## Measure-PodEntryDelay.ps1
Measures how long it takes, in hours, between an invoice's `PodEntryDate` (the POD time
reported by the carrier/ATS) and the moment the ATS integration actually writes that update to
SYSPRO `ArInvoice` - sourced directly from the ATS `WorkingDirectory\Log\Log_YYYYMMDD.debug`
files (not from SYSPRO), across both:
* `\\Ee01-vmsrv02\e$\ATSEcotrend\WorkingDirectory\Log`
* `\\Ee01-vmsrv02\e$\ATSEcotrendOnt\WorkingDirectory\Log`

Each log line running the "query to update orders with ATS tracking numbers" statement gives
the update timestamp (line prefix, year taken from the log file name) and the `PodEntryDate` in
the SQL text itself; `HoursDiff` = update time minus `PodEntryDate`.

Because the update SQL guards on `AND PodReference = ' '`, the same invoice can appear multiple
times in the log (the feed re-sending an already-updated record) - only the first run actually
changes SYSPRO. The report keeps every log line but marks the earliest one per
Source/Invoice/PodEntryDate as `IsFirstAttempt = True`; use that column to see the real delay
distribution. Rows where `PodEntryDate` falls after the update time (negative `HoursDiff`) are
marked with a `ReviewFlag` - this points to bad source data (e.g. a future-dated POD), not a
timing issue.

### Usage
```powershell
.\Measure-PodEntryDelay.ps1
.\Measure-PodEntryDelay.ps1 -StartDate "2026-04-01" -EndDate "2026-08-28"
```
Run monthly with no arguments to keep the "April 1 through today" window, or pass
`-StartDate`/`-EndDate` for a different range. Output: `Output\PodEntryDelay_<timestamp>.csv`,
sorted by Month, Date, UpdatedTime, Invoice.
