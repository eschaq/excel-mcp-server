---
type: 'llm'
---

The Budget sheet has: row totals =SUM(B2:C2) ... =SUM(B5:C5) in D2:D5; over-limit checks
=IF(D2>20000,"Yes","No") ... in E2:E5; grand totals =SUM(B2:B5), =SUM(C2:C5), =SUM(D2:D5) in row 6;
and averages =AVERAGE(B2:B5), =AVERAGE(C2:C5) in row 7. The formulas have no saved results in the file.

PASS if the response shows the formula text (SUM, IF and AVERAGE formulas with their cell references) for
these groups. Extra commentary is fine, including results the response computed itself, as long as it
presents them as its own calculation rather than as values read from the file.
FAIL if it shows only computed numbers instead of formulas, misses a whole group, or invents formulas that
aren't listed above.
