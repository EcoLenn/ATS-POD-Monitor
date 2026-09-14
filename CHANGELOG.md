2026-08-27
- Initial version
- Compare-PodData.ps1: compares carrier POD CSV against SYSPRO ArInvoice, with time zone
  conversion per account province (ON = Eastern, BC = Pacific)
- Generates PodComparison, PodUpdate (SQL), and PodUnmatched reports

2026-08-28
- Measure-PodEntryDelay.ps1: parses ATS WorkingDirectory\Log *.debug files from both
  ATSEcotrend and ATSEcotrendOnt servers to measure the delay between an invoice's
  PodEntryDate and when the ATS integration actually wrote it to ArInvoice
- Flags re-attempts on invoices already updated (IsFirstAttempt) and rows where
  PodEntryDate is after the update time (ReviewFlag)
