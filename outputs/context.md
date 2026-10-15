# Project context — LINE Reminder Bot

Last updated: 2026-10-05 (Asia/Bangkok)

## What this project is

A LINE bot built with Google Apps Script and Google Sheets. It helps users manage appointments, savings transactions, and installment contracts, with scheduled LINE reminders. It supports private chats and group chats.

## Current working files

- `LineReminderBot.gs` — current Apps Script implementation.
- `prompt_history.md` — older command notes; may not match the latest behavior.
- `memory.md` — persistent project snapshot and architectural decisions.
- `markdown.md` — setup and user/administrator operating guide.
- `bussiness_rule.md` — confirmed product rules.
- `history_prompt.md` — chronological summary of user requests.
- `suggestion.md` — recommendations and limitations.
- `errorlog.md` — reported incidents, diagnosis, and verification status.

## Current behavior

- Recognized commands start the bot without `hi bot`; `hi bot` remains an optional way to show the main menu. Unknown text in an inactive chat is ignored.
- Sessions are isolated by LINE conversation and expire after one minute without activity.
- Main command families: appointments (`C1`–`C3`), savings (`M1`–`M4`), installments (`L1`–`L2` plus submenu options), permissions (`A1` for a configured admin in personal chat), and exit (`X`).
- Appointments support optional time/location, LINE date/time picker, reminder choices (one day before, appointment date, or both), summaries, and deletion of open appointments dated today or later.
- Savings stores deposits and withdrawals. Balance is lifetime net; monthly deposit/withdrawal totals are shown for the current month; transaction history covers 15 days.
- Installments track contracts, principal, annual interest, separate principal/interest payments, and up to 16 months of transaction history. The user selects a contract before viewing history; the selection list only contains contracts with a transaction in that date window.
- Group records are shared within each group scope, and separated from private chats and other groups. Topic permissions can be configured for appointments, finance, and installments.
- Chinese-calendar reminders/image interpretation were cancelled and are out of scope.

## Runtime configuration

Apps Script Script Properties:

- `LINE_CHANNEL_ACCESS_TOKEN`
- `WEBHOOK_KEY`
- `BOT_ADMIN_USER_IDS` (comma-separated LINE user IDs)

The script timezone is `Asia/Bangkok`. `DAILY_HOUR` is currently `6`; `setup()` creates reminder and annual-interest daily triggers plus a one-minute idle-session trigger.

## Current task status and caveats

- The latest code edits include: accepting commands without a greeting, correctly preserving a zero principal balance, hiding fully settled contracts from outstanding results, grouping/formatting installment history, and limiting the history contract picker to contracts with records in the 16-month window.
- Those source edits require deploying a new Apps Script version to affect the live bot.
- No live LINE/Apps Script verification was performed for the latest edits. See `errorlog.md` for older unresolved reports and `markdown.md` for setup/troubleshooting.

## How to continue work

1. Read this file and `memory.md` for the project snapshot.
2. Read `bussiness_rule.md` before changing behavior.
3. Inspect `LineReminderBot.gs` as the implementation source of truth; do not rely on old menu details in `prompt_history.md` without checking the code.
4. After changes, update the relevant documentation and add newly reported errors to `errorlog.md` with their verification status.
