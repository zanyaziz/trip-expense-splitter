# Trip Expense Splitter

A free, no-signup, no-paywall alternative to Splitwise — as a single Excel file.
Log each trip expense once, mark who paid and who shares it, and the workbook
computes everyone's balance plus the shortest list of "X pays Y $Z" transfers
that settles everything.

Works in Microsoft Excel, Google Sheets, and LibreOffice. No macros, no
add-ins — plain formulas only.

## How to use (5 steps)

1. **People sheet** — type the traveler names (up to 12). Every other sheet
   updates automatically.
2. **Expenses sheet** — add one row per expense: date, description, who paid
   (dropdown), amount, then put `TRUE` in each traveler's column who shares
   that expense (dropdown). Leave blank anyone who doesn't share it.
3. **Share Each** is computed for you: `Amount ÷ number of TRUEs`
   (0 if nobody shares the expense).
4. **Balances sheet** — per person: Total Paid, Total Share Owed, and
   Net = Paid − Owed. Positive net means the group owes that person money;
   negative means they owe the group. The nets always sum to zero.
5. **Settle Up sheet** — the minimal transfer list. Each step matches the
   person who owes the most with the person who is owed the most, so at most
   11 transfers clear 12 travelers. A "Remaining unsettled" check at the
   bottom must read $0.00.

Tips: use one currency per trip. Unused traveler rows are ignored
automatically. Blank expense rows are skipped by every formula.

## How the settlement works

Classic greedy debt simplification. If Alice is owed $65, Bob owes $10 and
Carol owes $55, the sheet produces:

- Carol pays Alice $55.00
- Bob pays Alice $10.00

…instead of a tangle of pairwise IOUs. Each transfer clears at least one
person, so *n* travelers always settle in at most *n − 1* transfers.

## Why this exists

Expense-splitting apps keep adding paywalls for basic features. This template
does the core job — record, split, settle — in a file you own forever, free.

## License

MIT — do whatever you want with it. See [LICENSE](LICENSE).
