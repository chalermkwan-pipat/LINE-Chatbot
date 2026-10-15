# Project memory — LINE reminder and finance bot

Last updated: 2026-10-05 (Asia/Bangkok)

## Purpose

LINE bot for logging appointments, savings/deposits/withdrawals, and installment/loan balances. Data is stored in Google Sheets and reminders are sent through the LINE Messaging API. It supports one-to-one chats, groups, and rooms, with data scoped to the conversation.

## Source of truth

- Current implementation: `LineReminderBot.gs`
- Command/history notes: `prompt_history.md` (historical and may not reflect the latest menu)
- Current behavior and setup: `markdown.md`
- Confirmed product rules: `bussiness_rule.md`
- Older requests summarized: `history_prompt.md`
- Suggestions and known limitations: `suggestion.md`
- Errors and unresolved incidents: `errorlog.md`

When documents disagree, prefer the current script for implemented behavior and `bussiness_rule.md` for intended business behavior. Update these notes after future code changes.

## Current implementation snapshot

- Google Apps Script webhook handles LINE text messages, postbacks, quick replies, and date/time picker actions.
- `setup()` creates/extends `Tasks`, `Savings`, `Installments`, `Users`, and `Permissions` sheets, then installs daily reminder, daily annual-interest, and per-minute idle-session triggers.
- `DAILY_HOUR` is `6`; timezone is `Asia/Bangkok`.
- A recognized command can start a session directly; `hi bot` is still supported to show the main menu. Unknown chatter in an inactive chat is ignored.
- Sessions are scoped per chat and expire after 60 seconds without activity.
- The main menu includes appointment, finance, installment, and (for admins in personal chat) permission commands.
- Appointment create uses detail/location first, then a LINE date/time picker (time may be omitted), then reminder choice: one day before, on the appointment date, or both. Reminder messages label the scheduled appointment date, not the date the reminder was sent.
- Appointment cancellation/deletion lists open tasks dated today or later.
- Savings balance is lifetime net balance (all deposits minus all withdrawals); deposit and withdrawal totals shown alongside it are for the current month. Recent transaction history is 15 days.
- Installment principal and interest are tracked separately. Annual interest is recorded after each anniversary, using either original principal or balance at the anniversary according to the contract basis. Principal payments reduce principal; interest payments reduce interest due.
- Installment history covers 16 months. The user selects a contract first; the contract picker only includes contracts with at least one record in that window. History is grouped by contract and uses text-card formatting suitable for LINE.
- A contract with both principal balance and interest due at zero is omitted from outstanding balances. A contract with principal zero but interest due remains outstanding.
- Permission configuration is available only to a configured admin in personal chat; the admin can configure a whole group or an individual group member by topic.
- Chinese-calendar feature was explicitly cancelled and is not part of the product scope.

## Important architectural decisions

- `ownerKey` separates personal users (`user:<LINE user ID>`), groups (`group:<group ID>`), and rooms (`room:<room ID>`). Keep all reads/writes filtered by this key.
- Do not infer group admin status from LINE. The bot's admin list is configured in Script Properties as `BOT_ADMIN_USER_IDS`.
- Group permissions apply to appointment, finance, and installment topics. Group and member overrides are managed from the admin's personal chat.
- User-visible wording is Thai and generally ends with `ค่ะ`.

## Current setup properties

- `LINE_CHANNEL_ACCESS_TOKEN`
- `WEBHOOK_KEY` (web app URL uses `?key=...`)
- `BOT_ADMIN_USER_IDS` (comma-separated LINE user IDs)

See `markdown.md` for deployment and operating steps. Do not put actual secrets in these files.

## Outstanding verification

No live LINE/Apps Script deployment verification is implied by these documents. Recent code edits have not been exercised against the user's spreadsheet or LINE account. See `errorlog.md` for incidents and `suggestion.md` for recommended checks.
