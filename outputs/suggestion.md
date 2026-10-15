# Suggestions and implementation notes

Last updated: 2026-10-05

## Recommendations

- Keep `LineReminderBot.gs` as the implementation source of truth and update the six Markdown files whenever a user-visible flow or rule changes.
- Treat `prompt_history.md` as an older reference; it contains menu/start-flow details that have since changed. Prefer the current script and `markdown.md`.
- When installing updates, deploy a new Apps Script version and verify the LINE webhook still points to the active deployment.
- For incidents, capture timestamp, conversation type (private/group), command/flow, Apps Script execution result, and whether a LINE reply/push API error was logged. Do not store channel tokens or webhook secrets in logs.
- Consider adding a typed `YYYY-MM-DD` fallback for picker actions on LINE clients where the date picker quick reply is unavailable.
- Consider splitting long 16-month history into one message per contract or paginated output if LINE message length or readability becomes a problem.
- Consider a backup/export schedule for the Google Sheet before making schema changes or manually repairing old rows.

## UI note

LINE text messages do not render Markdown tables as true grid tables. The installment history uses grouped, bordered text cards to stay readable on mobile. A true structured card view would require a LINE Flex Message implementation.

## Data and permissions note

- The bot's data boundary is the conversation scope; anything recorded in a group is visible to members interacting with that group bot.
- The admin allowlist is a Script Property, not a LINE group role lookup. Configure IDs deliberately and keep them private.
- LINE profile/group-name endpoints can fail or return no name; show a fallback label rather than blocking permission setup.

## Product scope decisions

- Chinese calendar image parsing and recurring Chinese-calendar reminders were cancelled. Keep out of the active scope unless requested again.
- Cross-group/global read access was raised and then cancelled. Do not add a global-reader role based on that older request.
