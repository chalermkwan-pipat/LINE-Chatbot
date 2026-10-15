# Business rules — LINE reminder and finance bot

Last updated: 2026-10-05 (Asia/Bangkok)

## 1. Scope and data ownership

- Each LINE conversation has its own data scope. Private chat data belongs to that user; group and room data belong to that conversation.
- Group members share records created in their group. Different groups and private chats do not share data by default.
- Every stored record and lookup must be constrained by the conversation's `scopeKey`.
- A global/cross-group reader role is not in the current scope; the earlier request was cancelled.

## 2. Session behavior and commands

- A recognized command may activate the bot directly; `hi bot` is optional and opens the main menu.
- An unknown message in an inactive chat does not activate the bot.
- Latin command letters are accepted in either case.
- `X` ends use of the bot in the current chat. `/cancel` cancels the current interaction. Inactive sessions expire after 60 seconds without activity.
- Menus and results should be compact, easy to scan on a phone, in Thai, and use polite feminine wording (`ค่ะ`). Do not append a menu after every ordinary answer unless the flow explicitly requests one.

## 3. Appointments

- Users can create appointments with details and optional location/time. Unknown optional fields are left blank; do not fabricate values.
- Date/time selection is offered through LINE's picker. Date-only selection is valid.
- After creating an appointment, ask for one reminder policy: one day before, on the appointment date, or both. The default policy is on the appointment date when no different option is selected.
- Appointment summaries/reminders include the weekday and use Thai local dates (`Asia/Bangkok`). For advance reminders, display the scheduled appointment date, not the alert delivery date; each event carries its own date so mixed reminders remain unambiguous.
- Deletion choices include open appointments dated today or later. Deleted appointments are not active for summaries/reminders.
- Daily reminders are sent only for active/open tasks matching the configured reminder date and chats with notification enabled. Current trigger target hour is 06:00 Asia/Bangkok; Apps Script may execute with timing variation.

## 4. Savings and cash transactions

- A deposit or withdrawal must have a positive amount and a note. Amounts may have up to two decimal places.
- Transaction date is chosen through LINE's date picker; `วันนี้` means the current date in `Asia/Bangkok`.
- Balance equals all deposits minus all withdrawals across the life of that chat's ledger.
- Deposit and withdrawal cumulative subtotals are for the current calendar month only.
- Recent savings detail covers 15 days.

## 5. Installment contracts and debt

- Each contract has a title, creditor, debtor, starting principal, annual interest rate, note, interest basis, and unique `loanId`.
- Interest basis can be original principal or principal balance at the annual anniversary.
- Annual interest is calculated once per completed year/anniversary and recorded separately from principal. Calculated interest does not reduce principal.
- Paying principal reduces principal only. Payment cannot exceed the current principal balance.
- Paying interest reduces accrued unpaid interest only; it does not reduce principal.
- Transaction dates use the LINE date picker or the `วันนี้` quick action.
- If several contracts qualify for an action, ask the user to select a contract first.
- Outstanding view includes a contract while either principal balance or interest due is greater than zero. If both are zero, omit it and its latest-payment line.
- Installment history spans the current month and preceding 15 months (16 calendar months total), grouped by contract. Ask for the contract before showing history; offer only contracts with at least one transaction in that window.
- History is ordered newest first within each contract and should show transaction type/date, principal amount, interest amount, post-transaction principal, and note where present.

## 6. Permissions

- Permission topics are top-level only: appointments, finance, and installments.
- Only the configured admin (`BOT_ADMIN_USER_IDS`) sees/uses `A1`, and permission changes are made in the admin's personal chat.
- The admin can configure a whole group or a member of a group. Display group/member names when LINE provides them; use a fallback if not.
- For each topic, `Y` allows and `N` denies. Input is three comma-separated values in the order appointments, finance, installments.
- `9` restores default access for a group or makes a member inherit group settings.
- A configured admin's privileges do not create a different menu or automatic bypass inside a group; group/member permission rules apply there like for other users.

## 7. User experience and scope exclusions

- LINE messages must remain readable on mobile. Use concise headings, whitespace, icons, and grouped text cards; ordinary LINE text does not render Markdown tables as real grids.
- Chinese-calendar recurring reminders and image-based Chinese calendar parsing were cancelled and are excluded unless requested again.
- Calendar integrations with LINE group calendar are not implemented as a data sync feature; the bot currently manages its own Sheets-backed schedule and LINE notifications.
