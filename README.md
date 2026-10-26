# excel-agent

Agents and skills for spreadsheet work: auditing other people's models and building ones that survive review.

A spreadsheet is the most widely used programming environment in the
world and the one with the worst tooling. There is no test suite, no diff worth
reading, no type system, and the most expensive errors (a range that stopped
one row short, a hardcoded number pasted over a formula, a unit mismatch) look
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

These are skills for Claude. There is nothing to
build and nothing to import: you describe the work and the matching skill
fires on its own.

**In Claude Code**, clone the repository and link its skills and agents into the
folders Claude Code reads (`~/.claude/skills/` and `~/.claude/agents/`):

```bash
git clone https://github.com/Keremozdemirra/excel-agent ~/agent-packs/excel-agent
cd ~/agent-packs/excel-agent
mkdir -p ~/.claude/skills ~/.claude/agents
for s in skills/*; do ln -sfn "$PWD/$s" ~/.claude/skills/; done
for a in agents/*.md; do ln -sfn "$PWD/$a" ~/.claude/agents/; done
```

The links point back into the clone, so `git pull` there brings in whatever
ships next. To copy the files instead of linking them, replace the
two lines starting with `for` with:

```bash
cp -R skills/* ~/.claude/skills/
cp agents/*.md ~/.claude/agents/
```

Skills placed there are available in every session and every folder; start a
new session after installing.

Then simply ask. Each skill's `description` frontmatter is written to match how
the request actually gets phrased, in English or Turkish, so you do not name the
skill and generally should not have to think about which one applies.

**If nothing fires**, that is a defect in the skill rather than in how you
asked. The description was written for the wrong phrasing. Say what you asked
and what you expected, and it gets fixed. That feedback is more valuable than
working around it.

**What is actually built** is listed under Contents below and in the Done
section of [BACKLOG.md](BACKLOG.md). Everything under Queue is planned and does
not exist yet.

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

## Contents

| Skill | What it does |
| --- | --- |
| [model-audit](skills/model-audit) | Take apart an inherited workbook and rank what it gets wrong by how much of the answer each defect moves. |

| Agent | What it does |
| --- | --- |
| [veri-analisti](agents/veri-analisti.md) | Cleans and analyses CSV, Excel or JSON data and reports what it shows. |

Everything still under Queue in [BACKLOG.md](BACKLOG.md) does not exist
yet.
## Licence

MIT. See [LICENSE](LICENSE).
