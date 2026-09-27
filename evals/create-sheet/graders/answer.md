---
type: 'llm'
---

PASS if the response confirms a Returns sheet was added with headers Order ID, Reason, Refund Amount,
and that the workbook now has six sheets (Orders, Customers, Budget, Notes, Empty, Returns).
Mentioning only that the sheet and headers were added is also fine.
FAIL if it reports an error, a different sheet name, or different headers.
