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
> - Add a buffer of at least 2× the estimated row count so the user rarely runs out of space
>
> **After generating the file, tell the user:**
> - ✅ To add an expense: type into the next empty row — formulas already cover it
> - ✅ To edit an expense: click the cell and type the new value
> - ❌ To remove an expense: do NOT delete the row — select all cells in that row and press Delete to clear contents
> - ❌ Never insert or delete rows — this shifts Helper sheet references and breaks TOTAL OWED

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

## Adding, Editing, and Deleting Expense Rows

> ⚠️ **Never insert or delete rows.** The TOTAL OWED formulas reference a fixed range tied to the Helper sheet. Inserting or deleting rows shifts those references and breaks the totals. Always use the approaches below instead.

### Adding a new expense
Just type into the next empty row within the pre-built range — the Helper sheet already covers it. No other steps needed.

### Editing an expense
Click directly into the cell you want to change (expense name, amount, or share values) and type the new value. TOTAL OWED updates automatically.

### Removing an expense
**Do not delete the row.** Instead, select all cells in that row (columns A through the last person column) and press **Delete** to clear the contents. The row becomes blank and contributes $0 to all totals — exactly as if it never existed.

### If you've used up all rows in the buffer
Ask Claude to regenerate the file with a larger row count. When generating, ask for a buffer of at least 2× your expected expense count so this rarely comes up.

---

## Regenerating for a New Trip

Just paste the prompt above into a new Claude conversation. Claude will ask for names and row count, then generate a fresh file tailored to your group.
