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

## How to use this

These are skills for Claude, not a command-line tool. There is nothing to
install and nothing to import — you describe the work and the matching skill
fires on its own.

**In Claude Code or Cowork**, once the skills are on your machine:

```bash
bash ~/Desktop/agent/_setup/sync-skills.sh
```

That clones every agent repository and links its `skills/` into `~/.claude/skills`,
so they are available in every session and every folder. Re-run it whenever the
daily loop ships something new — it pulls rather than re-clones.

Then simply ask. Each skill's `description` frontmatter is written to match how
the request actually gets phrased, in English or Turkish, so you do not name the
skill and generally should not have to think about which one applies.

**If nothing fires**, that is a defect in the skill rather than in how you
asked. The description was written for the wrong phrasing. Say what you asked
and what you expected, and it gets fixed — that feedback is more valuable than
working around it.

**What is actually built** is the Done section of [BACKLOG.md](BACKLOG.md).
Everything under Queue is planned and does not exist yet. The daily loop builds
one item a day; the table above is the intended shape, not the current state.

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
