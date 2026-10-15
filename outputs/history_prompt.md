# Request history summary

This file summarizes the user's product requests in chronological themes. It is not a verbatim transcript; for exact messages refer to the conversation and `prompt_history.md`.

## Product foundation

- Build a LINE chat bot for reminders and work schedules, with multiple users and Google Sheets storage.
- Provide daily notifications, setup instructions, Apps Script source, `/help`, task summaries by date, and task deletion.
- Support group chats with data isolated per group and private chat.

## Chat flow and menus

- Initially use `hi bot` to start and `bye bot` to stop, then allow menu-based exit and automatic stop after one minute idle.
- Let users select the main topic before entering details; after each saved record, offer continue, return to main menu, or exit.
- Later request: allow known commands to start directly without requiring `hi bot`; keep `hi bot` as an optional menu shortcut.
- Make menus compact and scannable, use appointment/finance/installment headings and codes (`C*`, `M*`, `L*`, `A1`), and use `X` to exit.
- Accept lowercase and uppercase Latin command codes.
- Earlier requests to append menus after every answer were later cancelled; only show help/menu when requested or part of a flow.

## Appointments

- Save and delete appointments, summarize a selected date, and show today's/tomorrow's work.
- Permit blank appointment time or location while preserving the specified input structure.
- Ask after creation whether to notify one day before, on the appointment date, or both; default to the appointment date.
- Include weekday names in reminders and date summaries.
- Improve output formatting and allow date/time selection via LINE calendar picker, with date-only choice available.
- For deletion, list appointments from today onward.
- Chinese-calendar recurring feature and image recognition were discussed, then explicitly cancelled; do not reintroduce without a new request.

## Savings and money

- Track deposits and withdrawals, current balance, and transaction history for the most recent 15 days.
- Use current date for money transactions by default/quick selection; provide a calendar picker rather than requiring date typing.
- Show balance first and show cumulative deposit/withdrawal totals for the current month.
- Add shortcuts for deposit, withdrawal, balance, and savings menu.

## Installment contracts

- Track title, creditor, debtor, starting principal, annual interest rate, note, and interest calculation basis.
- Keep interest separate from principal; let user record principal and interest payments separately.
- Calculate annual interest at the anniversary using either original principal or balance basis; allow calculating from latest balance as a separate action.
- Choose the contract before posting a payment/calculation when more than one contract exists.
- Show outstanding principal, interest due, and last payment date; omit contracts with both balances at zero.
- Provide up to 16 months of history, grouped by contract and formatted for LINE. Choose the contract before showing history, and only offer contracts with records in the 16-month window.
- Recent requested behavior includes selecting the transaction date from LINE's picker or `วันนี้`.

## Permissions and scope

- Permission controls should apply by main topic (appointments, finance, installments), not by every submenu.
- Allow an admin to configure access in personal chat for an entire group or a member; group names/member names should be shown when available.
- Let ordinary members of a permitted group use that group's enabled functions; admin role in LINE should not create a different group menu.
- Admin command is visible only to configured admin account in personal chat.

## Documents and maintenance

- User has requested current scripts, troubleshooting, help/menu summaries, setup guide/PDF, and now a six-file Markdown documentation set.
- Keep documentation aligned with deployed source and record future issues in `errorlog.md`.
