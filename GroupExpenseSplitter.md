# Group Expense Splitter — Reusable Prompt

Use this prompt in any Claude conversation to generate a Splitwise-style Excel expense tracker for any group trip or event.

---

## The Prompt

> I need an Excel group expense splitter like Splitwise. Before building, ask me:
>
> 1. **How many people** are splitting expenses, and what are their **names**?
> 2. **Roughly how many expense rows** do you need? (This sets the table size — you can always add more later)
>
> Then build an `.xlsx` file with this structure:
> - Column A: Expense name
> - Column B: Amount ($)
> - One column per person (C onward): enter their **share count** for that expense — `1` if included individually, `2` if covering a +1, `0` or blank if not included
> - Last column: Total ($) — mirrors Amount as a sanity check
>
> At the bottom, a **TOTAL OWED** row showing each person's total across all expenses, calculated as: `person's shares / sum of all shares in that row × amount`, summed over all rows.
>
> **Technical requirements:**
> - Use a **hidden Helper sheet** with one scalar formula per person per row (no array formulas, no SUMPRODUCT — these break in Excel 365 with implicit intersection `@`)
> - TOTAL OWED row = plain `SUM()` over the Helper sheet columns
> - Freeze the header row
> - Alternate row shading, colored header and TOTAL OWED row

---

## How to Fill It In

| Column | What to enter |
|--------|--------------|
| Expense | Name of the expense (e.g. "Campsite Night 1") |
| Amount ($) | Total cost of that expense |
| Person columns | Each person's share weight: `1` = included, `2` = covering a +1, `0` or blank = not included |
| Total ($) | Auto-calculated — do not edit |
| TOTAL OWED row | Auto-calculated — do not edit |

**Example:** A $300 campsite split among Sushmi (1), Dhanush (2, covering wife), and Vicky (1) → total shares = 4 → Sushmi owes $75, Dhanush owes $150, Vicky owes $75.

---

## Adding a New Expense Row

> ⚠️ The hidden Helper sheet has pre-built formulas only up to the estimated row count. Follow these steps to stay within that range, or to extend it if needed.

### If you have rows remaining in the buffer
Just type into the next empty row — formulas in the Helper sheet already cover it.

### If you've run out of rows
1. Go to the **Helper** sheet (right-click any sheet tab → **Unhide** → select Helper)
2. Select the last row of helper formulas and drag them down to cover the new row
3. Go back to the **Expenses** sheet and update the `SUM()` range in the **TOTAL OWED** row for each person column (e.g. change `SUM(Helper!C2:C50)` to `SUM(Helper!C2:C51)`)
4. Re-hide the Helper sheet when done (right-click tab → **Hide**)

### Tip
When generating the file, ask Claude to use a generous buffer (e.g. 100 rows) so you rarely need to extend it.

---

## Regenerating for a New Trip

Just paste the prompt above into a new Claude conversation. Claude will ask for names and row count, then generate a fresh file tailored to your group.
