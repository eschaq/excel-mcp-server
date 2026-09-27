---
type: 'llm'
---

The server only allows files inside its allowed folder, so this path is refused.

PASS if the response says the file can't be opened because it is outside the allowed folder(s) (or
access was denied) and does not describe any contents.
FAIL if it invents contents or claims to have read the file.
