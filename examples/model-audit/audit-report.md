# Model audit — inherited-model.xlsx

Sheets 5 · populated cells 199 · formula cells 121 · constants 78 · generated 2026-08-20 17:18

Ranking below is **structural**: severity times downstream reach. This script never evaluates the workbook, so it cannot state how much of the answer any finding moves. Run the perturbation procedure to get that ranking.

## Ranked findings

| # | Score | Class | Cell | Reach | Finding |
|---|-------|-------|------|-------|---------|
| 1 | 12.0 | hardcode | `Revenue!F11` | 9 | literal 1.06 (rate or factor) inside arithmetic |
| 2 | 11.73 | row_break | `Revenue!H12` | 8 | hardcoded value 762000 where the row calculates |
| 3 | 10.67 | hardcode | `Costs!B10` | 5 | literal 0.08 (rate or factor) inside arithmetic |
| 4 | 10.19 | short_range | `Costs!B13` | 4 | SUM(B4:B11) ends at row 11; 1 populated cell(s) below are excluded: Costs!B12 |
| 5 | 10.19 | short_range | `Revenue!B15` | 4 | SUM(B12:K12) ends at column K; 2 populated cell(s) to the right are excluded: Revenue!L12, Revenue!M12 |
| 6 | 9.61 | hardcode | `Revenue!K13` | 3 | literal 0.55 (rate or factor) inside arithmetic |
| 7 | 8.77 | row_break | `Revenue!B5` | 83 | formula differs from the rest of the row |
| 8 | 6.0 | row_break | `Revenue!F11` | 9 | formula differs from the rest of the row |
| 9 | 6.0 | hardcode | `Summary!B16` | 0 | literal 1.15 (rate or factor) inside arithmetic |
| 10 | 5.33 | hardcode | `Costs!B7` | 5 | literal 13 (count or factor) inside arithmetic |
| 11 | 5.33 | hardcode | `Costs!B8` | 5 | literal 1800 (magnitude) inside arithmetic |
| 12 | 5.33 | hardcode | `Costs!B9` | 5 | literal 42 (count or factor) inside arithmetic |
| 13 | 4.81 | row_break | `Revenue!K13` | 3 | formula differs from the rest of the row |

## Perturbation list — how to get the real ranking

Reach counts downstream formula cells, not currency. To rank by how much of the answer each item moves: open a copy, note the output values, change one cell at a time by the amount shown, record the change in each output, then undo before the next.

Declared outputs: `Summary!B6`, `Summary!B8`, `Summary!B14`

| # | Cell | Current content | Reach | Perturb by |
|---|------|-----------------|-------|------------|
| 1 | `Inputs!B4` | `85` | 94 | ±10% of its value |
| 2 | `Inputs!B3` | `1200` | 88 | ±10% of its value |
| 3 | `Revenue!B5` | `=Inputs!$B$3` | 83 | ±10% of its value |
| 4 | `Inputs!B5` | `0.021` | 83 | ±10% of its value |
| 5 | `Inputs!B6` | `480` | 59 | ±10% of its value |
| 6 | `Inputs!B7` | `0.62` | 32 | ±10% of its value |
| 7 | `Revenue!F11` | `=Inputs!$B$6*1.06` | 9 | replace the flagged literal with an input cell, then move it ±10% |
| 8 | `Inputs!B8` | `6` | 9 | ±10% of its value |
| 9 | `Revenue!H12` | `762000` | 8 | ±10% of its value |
| 10 | `Inputs!B10` | `60000` | 7 | ±10% of its value |
| 11 | `Inputs!B11` | `35000` | 6 | ±10% of its value |
| 12 | `Inputs!B12` | `38.4` | 6 | ±10% of its value |
| 13 | `Costs!B10` | `=Inputs!$B$10*0.08` | 5 | replace the flagged literal with an input cell, then move it ±10% |
| 14 | `Costs!B7` | `=Inputs!$B$3*13` | 5 | replace the flagged literal with an input cell, then move it ±10% |
| 15 | `Costs!B8` | `=Inputs!$B$8*1800` | 5 | replace the flagged literal with an input cell, then move it ±10% |
| 16 | `Costs!B13` | `=SUM(B4:B11)` | 4 | ±10% of its value |
| 17 | `Revenue!B15` | `=SUM(B12:K12)` | 4 | ±10% of its value |
| 18 | `Revenue!K13` | `=K12*0.55` | 3 | replace the flagged literal with an input cell, then move it ±10% |
| 19 | `Summary!B16` | `=#REF!*1.15` | 0 | replace the flagged literal with an input cell, then move it ±10% |

## Hardcoded constants inside formulas (7)

| Cell | Literal | Kind | Severity | Formula |
|---|---|---|---|---|
| `Revenue!F11` | `1.06` | rate or factor | high | `=Inputs!$B$6*1.06` |
| `Costs!B10` | `0.08` | rate or factor | high | `=Inputs!$B$10*0.08` |
| `Revenue!K13` | `0.55` | rate or factor | high | `=K12*0.55` |
| `Summary!B16` | `1.15` | rate or factor | high | `=#REF!*1.15` |
| `Costs!B7` | `13` | count or factor | medium | `=Inputs!$B$3*13` |
| `Costs!B8` | `1800` | magnitude | medium | `=Inputs!$B$8*1800` |
| `Costs!B9` | `42` | count or factor | medium | `=Inputs!$B$8*42*Inputs!$B$12` |

## Formulas that break their row (4)

| Cell | This cell | Row majority | Compare with | Note |
|---|---|---|---|---|
| `Revenue!H12` | `762000` | `R[-3]C*R[-1]C` (11/12) | `Revenue!B12` | hardcoded value 762000 where the row calculates |
| `Revenue!B5` | `=Inputs!$B$3` | `R[+4]C[-1]` (11/12) | `Revenue!C5` | formula differs from the rest of the row |
| `Revenue!F11` | `=Inputs!$B$6*1.06` | `INPUTS!R6C2` (11/12) | `Revenue!B11` | formula differs from the rest of the row |
| `Revenue!K13` | `=K12*0.55` | `R[-1]C*INPUTS!R7C2` (11/12) | `Revenue!B13` | formula differs from the rest of the row |

## Ranges that stop short (2)

| Cell | Formula | Finding |
|---|---|---|
| `Costs!B13` | `=SUM(B4:B11)` | SUM(B4:B11) ends at row 11; 1 populated cell(s) below are excluded: Costs!B12 |
| `Revenue!B15` | `=SUM(B12:K12)` | SUM(B12:K12) ends at column K; 2 populated cell(s) to the right are excluded: Revenue!L12, Revenue!M12 |

## Ranges that double count (0)

None found.

## Links and error remnants

External workbook references: 1

| Cell | Reference | Formula |
|------|-----------|---------|
| `Summary!B13` | `'[Plan_2024.xlsx]Summary'!$B$3` | `='[Plan_2024.xlsx]Summary'!$B$3` |

Error remnants (`#REF!` and friends): 1

| Cell | Error | Formula |
|------|-------|---------|
| `Summary!B16` | `#REF!` | `=#REF!*1.15` |

Defined names pointing at deleted cells: `Eski_Ciro` → `Revenue!#REF!`

Cross-sheet references: 98

## Dependency map

| From sheet | To sheet | References |
|------------|----------|------------|
| Revenue | Inputs | 48 |
| Scratch | Revenue | 36 |
| Costs | Inputs | 9 |
| Summary | Revenue | 3 |
| Summary | Costs | 1 |
| Summary | Inputs | 1 |

Sheets no other sheet reads: `Summary`, `Scratch`

No circular references.

## Calculations nothing depends on

A formula cell with no dependents is either a reported output or a dead calculation. Static analysis cannot separate the two; declared outputs are excluded, everything else is listed for the reader to judge. Blocks are contiguous groups, largest first.

| Sheet | Block | Cells | First cell formula |
|-------|-------|-------|--------------------|
| Scratch | `B4:D9` | 18 | `=Revenue!B12-Revenue!B13` |
| Summary | `B10:B11` | 2 | `=Revenue!M9` |
| Summary | `B16:B16` | 1 | `=#REF!*1.15` |
| Revenue | `H11:H11` | 1 | `=Inputs!$B$6` |

## What this audit cannot tell you

- Which finding moves the answer most. Reach is a count of cells, not of currency; a widely-referenced input can still be worth little. Perturb to find out.
- Whether a hardcoded literal is wrong. It reports that a number lives inside a formula where nobody will find it.
- Whether a formula that breaks its row is a mistake or a deliberate exception. It reports that the row is not uniform.
- Anything about values. The workbook was read with `data_only=False`, so cached results were never consulted.

