---
name: model-audit
description: "Take apart a spreadsheet somebody else built and report what is wrong with its structure, ranked by how much of the answer each defect moves. Maps the dependency graph, finds numbers typed inside formulas, ranges that stop short of the block they sum, rows where one cell breaks the pattern, external links and #REF! remnants, and calculations nothing depends on. Use whenever the ask is to check, audit, inherit, take over, sanity-test or trust a workbook before acting on it: \"bu excel'i incele\", \"modeli denetle\", \"bu tabloda hata var mı\", \"bu dosyayı devraldım\", \"formülleri kontrol et\", \"bu toplam neden tutmuyor\", \"excel'e güvenebilir miyim\", \"audit this workbook\", \"check this spreadsheet for errors\", \"I inherited this model\", \"what is hardcoded in here\", \"why doesn't this total tie\", \"map the dependencies\". Use it before a board pack, a budget submission, a price decision or a handover. Not for building a model, not for cleaning data, not for reading values out of a workbook."
---

# Model Audit

A workbook you did not build is a claim, not a calculation. The claim is usually
accepted because the file opens, the numbers look plausible and the arithmetic is
invisible: a total that stops one row short of the block it sums looks exactly like a
total that does not, and a growth rate typed into the middle of a formula looks exactly
like a growth rate read from an input cell. Both survive every review that consists of
scrolling.

This skill reads the formulas rather than the values, which is where the defects live.
It produces two things a manual read cannot: a dependency graph, so you know which cells
carry the answer before you argue about any of them, and a normalised comparison of every
formula against its neighbours, so a single broken cell in a row of two hundred is found
in seconds rather than never.

## When to use / when not

Use it on any workbook you are about to rely on and did not write: an inherited budget, a
model from a bank or an agency, a file a departing colleague left behind, a supplier's
pricing sheet, or your own model six months later. Use it before the number leaves the
building. Use it when a total does not tie and nobody can say why.

Do not use it to build a model, to clean or reshape data, or to extract values — those are
different jobs. Do not use it on a workbook whose formulas have been stripped: a
values-only export has nothing to audit, and the finding is that the file is unauditable
in that form. Ask for the live file. Do not use it as a substitute for asking the author
what a cell means; the audit finds structural defects, not misunderstood intent.

The script reads `.xlsx` and `.xlsm`. For `.xls` or `.xlsb`, convert to `.xlsx` first and
say in the report that you audited a conversion.

## Method

**1. Work on a copy and never write to the original.** Copy the file to a working
directory before anything else. The audit itself only reads, but the perturbation step in
`5` edits cells, and a file you have edited is no longer evidence.

**2. Run the structural pass.** See below for the command. It reports the five classes:
hardcoded constants inside formulas, formulas that break their row, ranges that stop
short, links and error remnants, and calculations nothing depends on. Read the whole
report before forming a view; the classes interact, and a cell that appears in two of them
is where you start.

**3. Read the dependency map before you read any finding.** The sheet-to-sheet table tells
you the direction of the model. A sheet nothing reads is either the output or dead weight,
and knowing which changes how you read the rest. If the map shows circular references,
stop and resolve those first — every number downstream of a circular reference depends on
the iteration setting rather than on the model.

**4. Triage the findings into corrections and judgements.** A range that stops short, a
`#REF!` remnant and a typed number sitting where the row calculates are corrections: they
are wrong regardless of anyone's view, and they are not negotiable. A literal inside a
formula, an unusual formula in a row, and a link to an external workbook are judgements:
they may be deliberate. Ask the author about the judgements; fix the corrections. Keep
the two apart in the report, because a reader who cannot tell them apart will discount
both.

**5. Perturb to get the real ranking.** The script ranks by severity times reach — the
number of formula cells downstream of the defect. Reach is a proxy and nothing more. A
cell read by four hundred formulas can be worth nothing to the answer; a cell read by one
can be worth all of it. To find out what each defect is actually worth:

   a. Open the copy with automatic calculation on. Record the baseline value of every
      output cell you care about. Write those numbers down outside the file.
   b. Take the perturbation list in the order printed. Change one cell only.
      For an input constant, move it by an amount the model's owner would defend — the
      script prints ±10% as a placeholder, but use the real range where you know it.
      For a hardcoded literal, first replace the literal with a reference to a new input
      cell holding the same value, confirm the outputs have not moved, then move that
      input cell.
   c. Record the new value of every output. Compute the change in currency and in per
      cent.
   d. Undo. Confirm the outputs have returned to baseline before the next cell. A
      perturbation you forgot to undo contaminates every measurement after it.
   e. For corrections, measure once rather than perturbing: fix the short range, or the
      `#REF!`, and record what the fix does to each output. That delta is an error the
      workbook was carrying, not a sensitivity.
   f. Rank by the per cent change in the output that matters. That is the ranking the
      request asked for. The script's ranking is a starting order for this work, not a
      substitute for it.

   Two cautions. Deltas do not add: correcting a short range changes what a rate
   perturbation is worth, so compute the combined corrected case separately and label it
   as such. And a perturbation that moves nothing is a finding in its own right — a cell
   that carries no weight is one nobody needs to argue about.

**6. Report the ranking, not the list.** Lead with the corrections and what they are
worth. Name the two or three cells that carry the answer. Everything below roughly 1% of
the output goes in an appendix however wrong it is, and saying so is the service.

## Running the script

```bash
pip install openpyxl --break-system-packages   # if openpyxl is missing
# The relative path only resolves from the skill's own directory; locate the script directly.
MODEL_AUDIT="$(find "$HOME/.claude/skills" "$HOME/agents" -name audit.py -path '*model-audit*' -print -quit 2>/dev/null)"
python3 "$MODEL_AUDIT" model.xlsx \
    --outputs "Summary!B6,Summary!B8,Summary!B14" \
    --out audit-report.md \
    --json audit-findings.json
```

If `$MODEL_AUDIT` comes back empty the skill lives somewhere these roots do not
cover — ask the user where it is installed rather than guessing a path.

With no `--out` the report goes to stdout. `--json` is for chaining; the markdown is the
deliverable.

| Option | Effect |
|--------|--------|
| `--outputs` | Cells you treat as the model's answers. Excluded from the dead-calculation list and named in the perturbation procedure. Supply them; without them the dead-calculation section lists your outputs as suspects. |
| `--top` | Rows in the ranked table. Default 25. |
| `--min-run` | Shortest run of adjacent cells tested for row consistency. Default 4. Lower it for narrow models, raise it to cut noise. |
| `--share` | Fraction of a run that must share a pattern before the rest count as breaks. Default 0.6. |
| `--gap` | Blank cells tolerated when working out where a summed block really ends. Default 1. |
| `--strict` | Exit 2 when any high-severity finding exists. For use in a check that must fail. |
| `--max-cells` | Guard against loading something enormous. Default 400000. |

The workbook is opened read-only with `data_only=False`, so cached values are never
consulted and the file is never modified. Nothing leaves the machine.

## The hard parts

**The first column of a time series always breaks its row.** In a monthly build, January
reads an opening balance from the inputs sheet while February onwards read the previous
month. The script reports it, correctly, as a formula that differs from the rest of the
row. It is not a defect. Recognise the shape and say so in the report rather than passing
it through; an audit that cries wolf on the first row is not read to the end.

**Hardcoded is not the same as wrong.** A number inside a formula is a number nobody will
find when it needs changing. The finding is about where the number lives, not about its
value. Write it that way, or the author will defend the value and the point will be lost.

**A formula cell with no dependents is either an output or dead.** Static analysis cannot
tell the two apart, which is why `--outputs` exists. Where the audit reports a dead block
of any size, look for the sheet it used to feed; abandoned scenario blocks and deleted
summary sheets are the usual explanation.

**Reach both overstates and understates.** A cell inside an `IF` branch that never
executes counts as fully connected. A cell feeding a single output that the whole
business runs on counts as one. This is precisely why step `5` exists.

**External links stop the audit at the file boundary.** A formula reading
`'[Plan_2024.xlsx]Summary'!$B$3` is a number of unknown provenance. The audit reports it
and cannot go further. Either obtain the linked workbook and audit it too, or record the
value as an unverified input.

**Array formulas, shared formulas and merged cells.** openpyxl returns array formulas as
objects rather than strings; the script handles the common case, but a workbook that leans
on dynamic arrays deserves a manual pass. Merged cells report against their top-left
anchor.

## Output format

A single markdown file, `audit-report.md`, in this order: the ranked findings table, the
perturbation list, then one section per defect class, the dependency map, the dead
calculations, and a closing statement of what the audit could not determine. The script
produces all of it.

On top of the script's output, write a covering section of no more than eight lines at the
head of the file:

```markdown
# <workbook> — audit
Audited: <file, date received, from whom>   Outputs: <cells>   Auditor: <name>

## Finding
Two sentences. What the workbook is for, and whether its answer stands.

## Corrections — wrong regardless of view
| Cell | Defect | Effect on <output> | Fixed? |

## Judgements — ask the author
| Cell | Question | Worth if changed |
```

Fill the effect columns from step `5`. Leave them empty and the report is a list of
observations, which is what the reader already had.
