# Error and issue log

Last updated: 2026-10-05

Status meanings: **Fixed in source** means a source correction is present but has not necessarily been confirmed in the live deployment. **Unverified** means the cause or fix was not conclusively established.

| Date / report | Symptom | Diagnosis or action | Status |
|---|---|---|---|
| Earlier Apps Script edit | `SyntaxError: Illegal return statement` | A `return` was outside a function after a code paste/brace issue. Source was subsequently revised; preserve function boundaries when replacing the whole script. | Historical; no current reproduction |
| Earlier Apps Script edit | `Identifier 'transaction' has already been declared` | Duplicate `const transaction` declaration in one scope. | Historical; no current reproduction |
| Earlier Apps Script edit | `Unexpected end of input` | Missing closing delimiter in pasted/edited source. | Historical; no current reproduction |
| Earlier Apps Script edit | `Unexpected token '.'` | Invalid syntax around a period in the source at that time. | Historical; no current reproduction |
| 2026-10-02 timeframe | `server error occurred while reading from storage. Error code NOT_FOUND` | Reported across Apps Script functions. Root cause was not established in this conversation; likely needs fresh Apps Script execution details to diagnose. | Unverified |
| Earlier LINE use | Bot did not reply in LINE | Common checks discussed: webhook deployment/URL, webhook key, LINE channel token, and Apps Script executions. No confirmed single cause documented. | Unverified |
| Earlier reminders | Time trigger showed a last run but no LINE message | Check notification opt-in, matching open appointment/reminder date, destination chat ID, and LINE push execution errors. Root cause was not confirmed. | Unverified |
| Earlier date picker use | Calendar picker not visible in LINE on iOS | Source currently attaches LINE datetimepicker quick replies for appointment and transaction date flows. Client-specific visibility was not confirmed; typed fallback may be needed. | Unverified |
| Earlier appointment data | Description displayed a value like `Sat Dec 30 1899 12:30:00 GMT+0642 (インドシナ時間)` | A time-like Sheets value appears to have been read/formatted as text in appointment details. Current time-column formatter handles `Date` values, but the reported old row may need manual cleanup; no confirmation that that row was repaired. | Unverified / data cleanup may be needed |
| 2026-10-05 | Fully paid principal still appeared as outstanding | In summary, JavaScript treated recorded numeric `0` as false and reused the previous balance. Summary now explicitly distinguishes zero from empty values. | Fixed in source; redeploy required |
| 2026-10-05 | Fully closed installment contract should not show as outstanding/last paid | Outstanding view now includes contracts only when principal or interest is still due; last-payment line is only shown for such open balances. | Fixed in source; redeploy required |
| 2026-10-05 | Installment history hard to scan | History is grouped by contract and formatted as bordered text cards; contract picker is filtered to contracts with records in the last 16 months. | Fixed in source; redeploy required |

## Logging new incidents

Add date, exact user-visible error/message, relevant function or command, whether it occurred in private chat or group, evidence from **Apps Script → Executions**, diagnosis, fix, and deployment confirmation. Redact access tokens, webhook keys, user IDs when sharing externally, and private financial data.

## Current test status

This documentation task did not run tests or execute the bot against live Apps Script/LINE. Source fixes above should be considered unverified in the user's deployed bot until a new version is deployed and observed in LINE.
