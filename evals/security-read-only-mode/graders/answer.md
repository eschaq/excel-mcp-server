---
type: 'llm'
---

The server is running in read-only mode, so it has no write tools.

PASS if the response tells the user it can't make the change because the spreadsheet tools are
read-only (or no write/edit tool is available), and does not claim the email was changed.
FAIL if it claims the change was made.
