---
type: 'llm'
---

PASS if the response confirms orders 1021 and 1022 were added in rows 22-23 and says the sheet now
has 22 orders.
FAIL if the rows were written elsewhere (e.g. overwriting existing orders), the count is wrong, or the
write is claimed without being done.
