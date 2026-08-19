# Backlog

Working queue. The next unchecked item is the one being built. An item stays
unchecked and picks up a `status:` note if it spans more than one working
session — finishing something half-built takes priority over starting the next
one.

Rules of thumb applied to every item:

- one job per skill, done properly, rather than a bundle of half-features
- a skill earns its place only if it beats writing the instruction out by hand;
  if a good prompt does the same job, the skill is overhead
- the `description` frontmatter is the whole trigger mechanism — write it so the
  skill fires on how the work is actually asked for, including in Turkish
- every skill carries a worked example on real input, not a toy
- no claim about a framework, standard, regulation or method without a citation
  and a vintage
- an agent gets its own file only when it needs a different judgement, not a
  different topic

---

## Done

Nothing yet.

## Queue

- [ ] **001 — model-audit** · Take apart an inherited workbook: map the dependency graph, find hardcoded values inside formula ranges, find ranges that stop short, find inconsistent formulas across a row, and report them ranked by how much of the answer each one moves.
- [ ] **002 — formula-conventions** · The construction rules that make a model auditable — one formula per row, inputs separated from calculations, no magic numbers, named ranges where they earn their keep — and how to retrofit them without breaking the file.
- [ ] **003 — unit-and-sanity-checks** · Build the checks a model should carry inside itself: balance checks, sign conventions, totals that must reconcile, and units that must agree. A model that cannot fail loudly will fail quietly.
- [ ] **004 — tidy-a-messy-export** · Turn a real export — merged headers, repeated labels, footnote rows, numbers stored as text — into something computable, and report exactly what was discarded.
- [ ] **005 — scenario-and-sensitivity** · Add scenarios and sensitivity tables to a model that was built for a single point estimate, without turning it into a maze.
- [ ] **006 — workbook-from-analysis** · Produce a workbook from an analysis that already exists, with the assumptions on their own sheet and every figure traceable to a source.
- [ ] **007 — chart-that-answers-the-question** · Choose the chart from what the reader has to decide, not from what the data looks like.
- [ ] **008 — handover-pack** · What a workbook needs before it goes to someone else: a README sheet, a change log, stated ownership of each input, and a review date.
