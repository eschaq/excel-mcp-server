#!/usr/bin/env bash
# Generate a ~55 MB CSV in the workspace (not committed to the repo).
set -euo pipefail
awk 'BEGIN { print "Order ID,Date,Customer,Region,Product,Qty,Unit Price";
             for (i = 0; i < 1200000; i++) print "1001,2026-01-04,Acme Corp,West,Widget,12,2.50" }' > big_export.csv
