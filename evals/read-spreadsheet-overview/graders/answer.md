---
type: 'llm'
---

The workbook has five sheets: Orders (21 rows incl. header, 8 columns), Customers (9 rows, 5 columns),
Budget (7 rows, 5 columns), Notes (7 rows, 3 columns) and Empty (no data).

PASS if the response lists all five sheets with row and column counts consistent with these (counting
data rows without the header, e.g. 20 orders, is also fine) and says the Empty sheet has no data.
FAIL if any sheet is missing, any count is wrong, or it invents sheets or data.
