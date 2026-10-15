# LINE bot — operation and setup

Last updated: 2026-10-05 (Asia/Bangkok)

## Files

- `LineReminderBot.gs` — Apps Script source to deploy.
- `prompt_history.md` — older command summary; check against the current code before relying on menu details.
- `memory.md` — project snapshot.
- `bussiness_rule.md` — intended behavior and business rules.
- `history_prompt.md` — chronological request summary.
- `suggestion.md` — follow-up suggestions and known limitations.
- `errorlog.md` — reported errors and diagnostic status.

## Initial setup

1. Create/open the Google Sheet that will hold the bot data, then open **Extensions → Apps Script**.
2. Replace the editor contents with `LineReminderBot.gs` and save.
3. In **Project Settings → Script Properties**, set `LINE_CHANNEL_ACCESS_TOKEN`, `WEBHOOK_KEY`, and `BOT_ADMIN_USER_IDS` (comma-separated user IDs). Keep their values private.
4. Run `setup()` once from the Apps Script editor and approve the requested Google permissions. It creates the five data sheets and sets up triggers.
5. Deploy as a **Web app**, executing as the sheet owner and allowing access needed for LINE's webhook. Copy the deployment URL.
6. In LINE Developers Console, set the webhook URL to `<web-app-url>?key=<WEBHOOK_KEY>`, enable webhook use, and verify the webhook connection.
7. Enable **Allow bot to join group chats** if group use is needed, then add the official account to the group.
8. After every code update, create a new deployment version and point LINE to the active web app deployment.

`setup()` configures daily reminders and annual-interest accrual for 06:00 Asia/Bangkok, plus an idle-session cleanup every minute. Apps Script time triggers may run slightly later than the target time.

## Starting and using the bot

- Type a known command (for example `C2`, `M4`, `L2`, `A1` where authorized, or `/help`) to start directly. `hi bot` remains an optional shortcut to display the main menu.
- Commands are case-insensitive for Latin letters. Use `X` to leave. Sessions expire after one minute without activity.
- `/help` lists accessible command families; `/cancel` cancels the current flow.
- In a group, use the commands in that group. Its data is shared with members of that same group and separate from private chats and other groups.

## Command map

| Code | Function |
|---|---|
| `C1` | Appointment menu: create, summarize by date, or delete a future/current appointment |
| `C2` | Create an appointment |
| `C3` | Summarize appointments for a selected date |
| `M1` | Savings menu |
| `M2` | Deposit |
| `M3` | Withdraw |
| `M4` | Show savings balance |
| `L1` | Installment menu |
| `L2` | Outstanding installment principal and interest |
| `A1` | Configure group/member topic permissions; admin in personal chat only |
| `X` | End the bot session |

Menu numbering may evolve; the menu returned by the deployed `/help` and the current source take precedence.

## Appointment flow

1. Send `C2`.
2. Enter `รายละเอียด | สถานที่`; leave location empty after the separator when unknown.
3. Use the LINE quick-reply date/time picker. Choose date and time, or choose date only when no time is needed.
4. Choose notification: `1` one day before, `2` on the appointment date, `3` both.
5. The daily notification trigger sends tasks for the current day to chats opted in to notifications.

To remove an appointment, enter `C1`, choose delete, then select from open appointments dated today or later. Direct `ลบงาน <รหัส>` is also supported by the source.

## Savings flow

- `M2` / `M3`: enter amount and note, then select a date in the picker or tap `วันนี้`.
- Amount must be greater than zero and can have up to two decimal places.
- Balance is lifetime deposits minus lifetime withdrawals. Monthly deposit and withdrawal totals are shown separately.
- `M1` contains a 15-day transaction-history view.

## Installment flow

- `L1` opens the installment menu: set up a contract, pay principal, pay interest, view outstanding balance, view 16-month history, view interest due, or calculate interest from the latest balance.
- Contract setup records title, creditor, debtor, starting principal, annual rate, note, and interest basis.
- For history, choose the contract first. Only contracts with at least one transaction in the 16-month window are offered. Entries are grouped by contract.
- Principal and interest transactions have a selectable date; use the date picker or `วันนี้`.

## Permissions

1. The admin sends `A1` in their personal chat.
2. Choose a registered group by displayed group name.
3. Choose whole-group or individual-member permissions.
4. Set appointment, finance, and installment access using three comma-separated `Y`/`N` values (example `Y,N,Y`). For a member, `9` inherits group defaults; for a group, `9` restores default access.

The configured admin IDs are not inferred from LINE's group roles. An admin in a group does not automatically bypass that group's topic permissions.

## Data sheets

- `Tasks`: appointment rows, status, date/time/location, reminder mode, and chat scope.
- `Savings`: deposits/withdrawals, transaction date, amount, note, and chat scope.
- `Installments`: starting contracts, principal payments, annual interest records, interest payments, basis, and loan ID.
- `Users`: chat target, notification opt-in, session state, pending interaction, group name, and selected contract.
- `Permissions`: per-group and per-member access by topic.

Use `scopeKey` to keep records isolated. Do not manually change IDs or scope keys unless performing a planned data repair.

## Troubleshooting checklist

- No reply: verify active deployment URL, `WEBHOOK_KEY`, LINE webhook setting, channel token, and Apps Script execution errors.
- Trigger ran but no reminder: check notification opt-in, appointment status/date/reminder mode, chat scope, target ID, and LINE push errors in execution logs.
- Picker absent: check that LINE quick replies are delivered and test on the target LINE iOS version; a typed date fallback may be needed if that client does not expose the picker.
- Permission issue: verify `BOT_ADMIN_USER_IDS`, then configure in personal chat and check the group/member row in `Permissions`.
- Read `errorlog.md` before treating a previously reported issue as resolved.
