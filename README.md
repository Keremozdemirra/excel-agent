# excel-agent

Agents and skills for spreadsheet work — auditing other people's models and building ones that survive review.

A spreadsheet is the most widely used programming environment in the
world and the one with the worst tooling. There is no test suite, no diff worth
reading, no type system, and the most expensive errors — a range that stopped
one row short, a hardcoded number pasted over a formula, a unit mismatch — look
exactly like correct output.

This repository holds the agents and skills for two jobs: taking apart a model
someone else built and finding what is wrong with it, and building one that
somebody else could take apart without finding anything.

## What this is not

It is not a spreadsheet library. Where a file has to be produced,
the skills call one; the value here is the judgement, not the writer.

It does not do financial analysis. Deciding what the model should say is
`finance-agent`'s job; this is about whether the model says what it claims to.

It will not rescue a workbook whose logic is wrong. It finds errors of
construction, not errors of thinking.

## Layout

```
agents/
  <name>.md           one specialist, its brief and its boundaries
skills/
  <name>/
    SKILL.md          the instruction, with triggering description frontmatter
    scripts/          only where deterministic code beats instruction
examples/
  <name>/             worked example on real input, with the output committed
```

## Roadmap

See [BACKLOG.md](BACKLOG.md). The first unchecked item is the one being built.

## Planned contents

Nothing here is built yet. This table is the intended shape, and the daily loop
fills it in one item at a time.

| # | Skill | What it does |
| --- | --- | --- |
| 001 | [model-audit](skills/model-audit) | Take apart an inherited workbook: map the dependency graph, find hardcoded values inside formula ranges, find ranges that stop short, find inconsistent formulas across a row, and report them ranked by how much of the answer each one moves. |
| 002 | [formula-conventions](skills/formula-conventions) | The construction rules that make a model auditable — one formula per row, inputs separated from calculations, no magic numbers, named ranges where they earn their keep — and how to retrofit them without breaking the file. |
| 003 | [unit-and-sanity-checks](skills/unit-and-sanity-checks) | Build the checks a model should carry inside itself: balance checks, sign conventions, totals that must reconcile, and units that must agree. |
| 004 | [tidy-a-messy-export](skills/tidy-a-messy-export) | Turn a real export — merged headers, repeated labels, footnote rows, numbers stored as text — into something computable, and report exactly what was discarded. |
| 005 | [scenario-and-sensitivity](skills/scenario-and-sensitivity) | Add scenarios and sensitivity tables to a model that was built for a single point estimate, without turning it into a maze. |
| 006 | [workbook-from-analysis](skills/workbook-from-analysis) | Produce a workbook from an analysis that already exists, with the assumptions on their own sheet and every figure traceable to a source. |
| 007 | [chart-that-answers-the-question](skills/chart-that-answers-the-question) | Choose the chart from what the reader has to decide, not from what the data looks like. |
| 008 | [handover-pack](skills/handover-pack) | What a workbook needs before it goes to someone else: a README sheet, a change log, stated ownership of each input, and a review date. |

## Licence

MIT. See [LICENSE](LICENSE).
