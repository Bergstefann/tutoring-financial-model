# Brief: Tutoring business financial model (Excel)

## What this is and why

A portfolio artefact for the **data/analytics pillar** of a graduate job search in
Flanders. The target market is Microsoft-stack, so this is Excel, not Google Sheets —
that decision is made, don't revisit it.

The subject is my own private tutoring business, which winds up when the school term ends
(~18 September 2026). The data is real: invoices, students, and lesson history already
live in the SQLite database that backs the Invoicer project
(`automated-invoicing-system`). This model is built on that, not on invented numbers.

The point of the artefact is to demonstrate real financial modelling discipline —
driver-based structure, documented assumptions, scenario and sensitivity analysis — on a
business I actually ran. Not spreadsheet decoration.

## Hard rule: anonymise before anything else

This goes in a **public GitHub repo**. The source data contains real students (minors),
their parents, and home suburbs.

- Every student becomes a stable pseudonym (`Student 01`, or a consistent fake given name).
  The mapping is generated once, kept **out** of the repo, and gitignored.
- No parent names, no email addresses, no street addresses. Suburb-level geography only,
  and only because travel cost genuinely depends on it.
- Rates and revenue figures are mine, not client data — those can stay real.
- Before any commit, grep the workbook and every intermediate file for real names and
  addresses. `openpyxl` files hold text in a shared strings table; check the extracted
  content, not just the cells you wrote.

If you can't anonymise something without destroying the analysis, stop and ask.

## Phase 1 — Inventory the source data. Build nothing yet.

Read the Invoicer SQLite schema and report what actually exists:

- Which tables and columns hold students, lessons, rates, invoices, dates, durations.
- Date range covered, number of students, number of lessons, number of invoice runs.
- Whether per-student rates vary, and whether rate changes over time are recorded.
- **Whether suburb / location is stored at all.** I believe it may not be. If it isn't,
  say so — do not infer it, and do not invent it.
- Any data quality problems: the known double-billing incident forked a duplicate student
  record, so check whether duplicate or merged student identities are still visible in the
  data and how they'd distort per-student revenue.

**CHECKPOINT — report this before modelling.** I'll fill in anything that has to come from
my own knowledge (suburb per student, travel distances, vehicle running cost).

## Phase 2 — The workbook

### Structure

Separate sheets, in this order. Inputs are never mixed into calculations.

1. **README / Legend** — what the model does, which cells are inputs, colour key, the
   as-at date, and a plain statement that student names are pseudonymised.
2. **Assumptions** — every input in its own labelled cell: hourly rates, super rate
   (12.5%), income tax provisioning rate, vehicle cost per km, working weeks per year,
   maximum teachable hours per week, travel time per suburb hop. Each with a source note
   saying whether it came from the data, from ATO/statutory rates, or from my own estimate.
3. **Data** — the anonymised lesson and invoice history, exported from SQLite as a proper
   table with typed columns. This is the factual base; nothing here is hand-edited.
4. **Revenue by student** — per student: lessons delivered, hours, effective hourly rate,
   gross revenue, share of total.
5. **Cost to serve** — travel cost per student (distance × cost per km × trips), plus
   travel *time* treated as unpaid capacity consumed. Produces contribution per student
   and contribution per hour committed, which is the number that actually ranks students.
6. **Capacity** — teachable hours per week, less travel time, giving a hard ceiling on
   students and on revenue. Show how far actual utilisation sat below the ceiling.
7. **P&L and provisioning** — gross revenue, costs, super at 12.5%, tax provision, net.
8. **Scenarios** — three named cases (base / more students at current rates / higher rate
   with fewer students), driven off the Assumptions sheet via a single scenario switch.
   No copy-pasted parallel models.
9. **Sensitivity** — a grid over two drivers (rate × utilisation) against net income.
   Build it as an explicit formula grid, not a native Excel Data Table.
10. **Charts** — revenue concentration by student, contribution per hour by suburb,
    utilisation against ceiling.

### Modelling conventions

- **Formulas, never hardcoded results.** `=SUM(...)`, not a Python-computed total pasted in.
- Every assumption referenced by cell (`=B5*(1+$B$6)`), never a magic number inside a
  formula.
- Colour convention: blue text for hardcoded inputs, black for formulas, green for
  cross-sheet links, yellow fill on key assumptions.
- Currency `$#,##0`, negatives in parentheses, percentages stored as fractions and
  formatted `0.0%`, years as text.
- Guard every denominator that can be zero.
- Formulas consistent across a row — a single edited cell mid-row is the commonest silent
  error in a model like this.
- **Stick to Excel-2007-era functions**: `SUMIFS`, `INDEX`/`MATCH`, `IFERROR`,
  `SUMPRODUCT`. Avoid `XLOOKUP`, `FILTER`, `UNIQUE`, `SORT` — they write badly from
  openpyxl and can silently produce truncated results.
- Verify by opening and checking real values, not by assuming the formula is right. A
  workbook with zero formula errors can still be entirely wrong.

## Phase 3 — Repository packaging

New repo, or a `spreadsheets/` area — recommend which and why, given the three-pillar
structure. Either way it sits **under the data pillar**, not as a separate weaker item.

- README with **screenshots** of the key sheets and charts. Nobody opens an `.xlsx` from
  GitHub, so the README has to carry the whole story.
- The README states the analytical findings, not just the mechanics: revenue
  concentration, which suburbs were unprofitable once travel was counted, how close to the
  capacity ceiling the business ran, what the sensitivity says about rate vs volume.
- State the anonymisation plainly. Handling real client data responsibly is itself worth
  saying out loud in a European job market.
- Note the as-at date and that the business wound up in September 2026.

## Standing rules

- **No `Co-Authored-By` trailers on any commit.** Ever.
- Conventional commits.
- Don't fabricate confidence. If travel data doesn't exist and I supply estimates, the
  workbook says they're estimates, in the cell.
- Anything irreversible or needing my login, hand back to me.
