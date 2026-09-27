---
type: 'llm'
---

Qty: 19 values (1 missing), min 2, max 100, mean about 26.21, median 14, stdev about 27.69.
Unit Price: 20 values, min 2.50, max 45.00, mean about 18.45, median about 13.62, stdev about 17.11.

PASS if the response gives all five statistics for both columns matching these values (reasonable
rounding is fine).
FAIL if any statistic is wrong or missing, or it treats the missing Qty as zero.
