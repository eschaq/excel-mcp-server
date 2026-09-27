---
type: 'llm'
---

Exactly four orders match: 1001 (Acme Corp, Widget, 2.50), 1007 (Hooli, Gizmo, 7.25),
1012 (Hooli, Widget, 2.50) and 1017 (Acme Corp, Gizmo, 7.25).

PASS if the response lists exactly these four orders.
FAIL if one is missing or an order from another region or at $10 or more is included.
