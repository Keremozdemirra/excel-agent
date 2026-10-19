# Worked example — an inherited FY2025 plan

## The file

`inherited-model.xlsx` is a twelve-month subscription plan across five sheets:
`Inputs`, `Revenue`, `Costs`, `Summary`, `Scratch`. 199 populated cells, 121 of them
formulas. `make-inherited-model.py` rebuilds it. The rebuilt file is not byte-identical,
because the workbook records when it was created, but its cells are, and the audit
reproduces every figure below: the same 13 of 13 findings.

Thirteen defects were planted, of the kinds that survive a real handover. Nobody was told
where they are before the audit ran.

## Running it

```bash
python3 make-inherited-model.py
python3 ../../skills/model-audit/scripts/audit.py inherited-model.xlsx \
    --outputs "Summary!B6,Summary!B8,Summary!B14" \
    --out audit-report.md --json audit-findings.json
```

Output: `audit-report.md`, `audit-findings.json`.

## What was planted, and what the audit found

| Planted | Found | Reported as |
|---------|-------|-------------|
| `Revenue!B15` totals Jan–Oct only | yes | short_range, `SUM(B12:K12)` excludes `L12`, `M12` |
| `Costs!B13` totals eight of nine cost lines | yes | short_range, excludes `Costs!B12` |
| `Revenue!H12` overwritten with the number 762000 | yes | row_break, hardcoded value where the row calculates |
| `Revenue!F11` carries a 6% price rise inside the formula | yes | hardcode (1.06) and row_break |
| `Revenue!K13` uses 0.55 rather than the margin input | yes | hardcode (0.55) and row_break |
| `Costs!B7` per-customer hosting cost of 13 typed in | yes | hardcode (13, count or factor) |
| `Costs!B8` per-seat tooling cost of 1800 typed in | yes | hardcode (1800, magnitude) |
| `Costs!B9` licence price of EUR 42 typed in | yes | hardcode (42, count or factor) |
| `Costs!B10` travel at 8% of marketing, typed in | yes | hardcode (0.08) |
| `Summary!B13` links to `Plan_2024.xlsx`, which nobody has | yes | external reference |
| `Summary!B16` survives a deleted sheet | yes | `#REF!` remnant, plus hardcode 1.15 |
| Defined name `Eski_Ciro` points at `Revenue!#REF!` | yes | listed under error remnants |
| `Scratch!B4:D9`, an abandoned upside case | yes | dead block, 18 cells |

All thirteen were reported. Nothing was missed.

## What the audit found that was not planted

**`Revenue!H11` is dead.** The August ARPU cell still calculates, but August revenue was
overwritten with a typed number, so nothing reads it. This is the second-order effect of
the hardcode and is exactly the kind of thing scrolling does not catch.

**`Summary!B10:B11` has no dependents.** Ending customers and revenue per head are
reported but feed nothing. They are outputs the auditor did not declare, not defects — a
reminder to pass every real output to `--outputs`.

**`Summary` and `Scratch` are read by no other sheet.** True and expected for `Summary`.
For `Scratch` it is the finding.

## The false positive, named

`Revenue!B5` is reported as breaking its row. It reads the opening balance from `Inputs`
while February to December read the previous month's closing figure. That is how a
monthly build is supposed to work. Every time series produces this finding in its first
column; recognise it and say so rather than passing it through.

## Structural ranking versus the real one

The script ranked by severity times reach. The perturbation procedure in step 5 of the
skill was then run against `Summary!B6` (EBITDA) and `Summary!B8` (EBITDA margin).
Baseline as received: revenue 7,260,239; EBITDA 219,813; margin 3.03%.

| Correction | Δ revenue | Δ EBITDA | Script rank | True rank |
|------------|-----------|----------|-------------|-----------|
| `Costs!B13` extended to `B12` | 0.00% | −65.51% | 4 | 1 |
| `Revenue!K13` restored to the margin input | 0.00% | +27.41% | 6 | 2 |
| `Revenue!F11` price rise moved to an input | −0.61% | −12.51% | 1 | 3 |
| `Revenue!H12` restored to `=H9*H11` | +0.38% | +7.74% | 2 | 4 |
| `Revenue!B15` extended to `M12` | +24.02% | 0.00% | 5 | 5 |
| `Costs!B10` travel rate ±10% | 0.00% | ±2.62% | 3 | 6 |

Three things this table is here to show.

**The structural ranking is not the real ranking.** The script's top finding is third on
the money. Its fourth is first. A 12,000 TRY line item left out of a cost total is worth
two thirds of EBITDA because EBITDA is a small difference between two large numbers —
nothing about the formula's structure says so. This is why the skill refuses to claim the
script ranks by value.

**A defect can move one output and not another.** Extending `Revenue!B15` changes revenue
by 24% and EBITDA by nothing, because EBITDA is built from the gross profit row rather
than from the revenue total. It still changes the margin, from 3.03% to 2.44%, because
the margin divides by the broken total. Rank against the output that matters, and say
which one you ranked against.

**Here the deltas happen to add.** Individually they sum to −94,249; taken together they
are −94,250. That is a property of this model, whose corrections touch separate additive
terms. Do not assume it. Compute the combined case separately, as was done here.

## The combined corrected case

All five corrections together: revenue 8,987,234 (+23.79%), EBITDA 125,563 (−42.88%),
margin 1.40% against the 3.03% the workbook reported. The plan as received overstates
EBITDA by 75% and overstates the margin by more than double.

`Summary!B14`, growth against prior year, could not be measured at all: it divides by
`Summary!B13`, which reads a workbook nobody has. The audit stops at the file boundary
and the report says so.
