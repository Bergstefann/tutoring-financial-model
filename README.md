# Tutoring Business Financial Model

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A driver-based Excel financial model of a private music-tutoring business, built from its own real lesson schedule and billing data for one school term. The annual figures here are an extrapolated run-rate, not a real full year since the business ceased services in September.

## Purpose

This is the data/analytics entry in my multi-project portfolio for my graduate job search in Flanders. The market runs on Microsoft-stack tools and Excel is used frequently to build driver-based structure, documented assumptions and conduct scenario and sensitivity analysis. I decided to apply this to the case of my own private practice to become more familiar with the industry standard toolset.

## Source of the data

The source of the data for this project is the Google Sheet "Lesson Schedule", used to manage the business day to day. It comprises a lesson attendance roll and a student config table exported into an anonymised, long-format CSV. Every derived figure in the workbook is an Excel formula.

## Findings

**P&L (9-week extract):** $11,840 gross revenue → $10,958 after $882 of allocated travel cost (7.4% of revenue) → $7,671 net income after 12.5% super and 20% tax provisioning. Extrapolated to a full year at the current pace (36 working weeks): ~$47,360 gross, ~$30,682 net, a run-rate rather than an actual annual figure.

**Scenarios and sensitivity: Rate is the lower-risk lever, while volume yields higher returns.** Max student volume (each timeslot in the week filled) (+30% volume, $40,628 net) modestly beats "higher rate, fewer takers" (+25% rate / −15% volume, $32,754) against the $30,682 base. Rate bumps are a low-risk effective lever until they begin to result in less customers. Volume is an effective lever in this scenario because allocated travel cost is constant regardless of lesson volume. This is because there are a set amount of campus visits a tutor can feasibly manage in a week, meaning extra volume must come from optimising the existing travel schedule. Each new 30 minute lesson is worth an extra $40 / week. At the current volume, bumping the rate per lesson by $1.22 increases gross income by $40 / week.

## Anonymisation

The real business data has been anonymised to Student 01–Student 37 format. No real name, parent name, email, or address appears anywhere in the repository or the workbook. This was double checked by grepping the built `.xlsx`'s shared-strings table, since that table is where leftover sensitive text tends to survive in `openpyxl` files. The pseudonym mapping lives in a local, gitignored file that's never committed. Rates and revenue are the real figures.

## Try it

Requires Python 3.13 with `openpyxl`, `pywin32`, `pymupdf`, and `pillow`.

```
scripts/extract.py             reads the raw export, anonymises, writes CSV
scripts/build_workbook.py      builds the workbook from the anonymised CSV
scripts/recalc.py              forces Excel to recalculate, checks for formula errors
scripts/screenshot_workbook.py exports the sheet screenshots used in this README
```

`data/` never leaves the machine it was built on; nothing under it is tracked by git.

## Known limitations

- **No suburb or location data exists in the source.** Travel is a flat weekly estimate (1.5 hrs/day, $98/week vehicle running cost) allocated by each student's share of lessons, not a distance calculation, so the model can't say which suburbs were unprofitable once travel is counted.
- **Only one 9-week term of data.** This model has nothing to say about trends over time.
- **The 20% income tax provisioning rate is a placeholder estimate and not a derived figure.** The business wound up mid financial year, so the real liability depends on total income across the whole year, which this model doesn't have.

## License

[MIT](LICENSE)
