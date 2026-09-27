---
type: 'llm'
---

PASS if the response confirms F1 = "Account Manager" (text), G5 = 42 (a number) and H2 = 2026-07-01
(a date), and that reading the sheet back showed these values.
FAIL if any value is wrong, the date was stored as some other date, or it claims success without
having read the data back.
