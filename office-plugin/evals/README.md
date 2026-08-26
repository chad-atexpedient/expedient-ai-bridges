# Agent eval harness

Grades CTRL agent runs on **resulting workbook state**, never on what the agent
claimed it did in the transcript.

## Why this exists

An agent can fix every bug it was asked to fix and still wreck the workbook.
The common failure is replacing a healthy formula with its computed value:
the numbers still look right today, and the model is dead. Transcript-based
review scores that run as a success. So does a passing unit suite.

This harness makes that failure visible and gradeable.

## Two-level scoring

| Score | Question | Fails when |
| --- | --- | --- |
| **Modif.** | Did the agent make the intended changes? | A target fix is missing or wrong |
| **Acc.** | Is the whole workbook still correct? | Anything else changed, including collateral damage |

Report both. `Modif.` alone rewards destructive runs. A run with
`Modif.=PASS, Acc.=FAIL` is the signal this harness was built to surface.

## The four checks

1. **`cells_match`** - graded output cells match the expected-value oracle,
   with relative tolerance for nonzero numbers and absolute tolerance at zero.
   Booleans must stay booleans.
2. **`target_fixes`** - intended edits match at the formula level, using
   quote-aware normalization: case and whitespace are ignored *outside* string
   literals, so `="Past Due"` is never conflated with `="PastDue"`.
3. **`no_mutation`** - protected sheets are untouched in both values and
   formulas, diffed over the union of seed and result cells.
4. **`unintended_edited_cells`** - first-class destructive-edit metric. Diffs
   every cell outside the intended set over the union of addresses, so
   deletions are caught even when the used range shrinks. Formula-replaced-by-value
   is called out explicitly.

## Hidden variants

Each task defines variants that perturb input literals. The correct *formulas*
stay identical across variants while the correct *values* change, so an agent
that memorized or hardcoded an answer fails as soon as the variant shifts.
Always record which variant a run used, and track pass rate per variant.

## Usage

```bash
node evals/run.mjs --list                                  # discover tasks
node evals/run.mjs --self-test                             # verify the grader itself
node evals/run.mjs --validate                              # verify every task spec is sound
node evals/run.mjs --check-corpus                          # verify every task discriminates
node evals/run.mjs --task margin-model-doctor --seed-only --out seed.xlsx
node evals/run.mjs --task margin-model-doctor --result after.xlsx
node evals/run.mjs --task margin-model-doctor --result after.xlsx --json
node evals/run.mjs --task margin-model-doctor --variant b --result after.xlsx
```

Exit code is `0` only when `Acc.` passes, so the runner drops straight into CI.

`npm run eval:self-test` is wired into `scripts/release-gates.mjs`. If the
grader ever stops catching the destructive run, the release gate fails. A
grader that silently degrades is worse than no grader.

## Writing a task

Tasks are data, not code. See `tasks/margin-model-doctor.json`.

| Field | Purpose |
| --- | --- |
| `prompt` | Exact text given to the agent |
| `seed.sheets` | Fixture workbook definition; `=FORMULA` strings become real formulas |
| `targetFixes` | Formula-level intended edits (drives `Modif.`) |
| `expectedValues` | Value-level oracle for computed results |
| `protectedSheets` | Sheets that must not change at all |
| `allowedEdits` | Cells that may change without counting as destructive |
| `variants` | Literal perturbations that defeat memorized answers |
| `budget` | Max tool calls and minutes per task |
| `referenceSolution` | The correct edit set, used by `--check-corpus` |
| `referenceDamage` | Collateral damage a careless run would cause, used to prove the task discriminates |

Seeds are **generated, never committed as binaries**, so the corpus stays
diffable and license-clean.

## Corpus policy

**The task set is the asset. The harness stays thin.**

Tasks derived from licensed or proprietary material must live in a private
corpus outside this repository. Before any workbook is used as a seed, scrub
it: hidden sheets, personal metadata, custom properties, comments, external
links, cached formula values, VBA, and non-builtin defined names all leak
answers or confidential data. External links pointing at an answer-key file
are a real and repeatedly observed leak.

Nothing in `evals/corpus/` is committed except this policy.

## Task-spec validation

A malformed task is worse than a missing one: it is a gate that reports success
while testing nothing. `--validate` refuses four traps, each confirmed against
the live grader before the validator existed:

| Trap | Symptom without validation |
| --- | --- |
| No `targetFixes` and no `expectedValues` | Every run scores `Modif.=PASS, Acc.=PASS` |
| A `targetFixes` entry the seed already satisfies | Task passes with the agent doing nothing |
| `protectedSheets` naming a nonexistent sheet | Protection is silently inert |
| `expectedValues` on a formula cell | Unsatisfiable: generated seeds carry no cached values |

It also requires at least two variants, at least one variant that actually
perturbs the seed, single-cell addresses, and references that resolve to real
seed sheets.

## Corpus discrimination

Validation proves a task is well-formed. It does not prove the task can tell a
good run from a bad one. `--check-corpus` does that, and it is the check worth
trusting.

Every task declares two reference runs:

| Field | Meaning |
| --- | --- |
| `referenceSolution` | The correct set of edits |
| `referenceDamage` | Additional edits a careless agent would make to healthy cells |

For each task and each variant, the harness asserts:

- `referenceSolution` alone scores `Modif.=PASS, Acc.=PASS`
- `referenceSolution` plus `referenceDamage` scores `Modif.=PASS, Acc.=FAIL`

That second assertion is the entire thesis. It was verified by sabotage:
removing the unintended-edit term from the `Acc.` calculation fails all 14
checks with the diagnosis "grader failed to flag the destructive run". Both
`--validate` and `--check-corpus` run in `scripts/release-gates.mjs` and CI.

## Current corpus

| Task | Failure mode it measures |
| --- | --- |
| `margin-model-doctor` | Hardcoded formula, wrong row reference, short SUM range |
| `quarterly-rollup` | Additive extension without disturbing prior periods |
| `lookup-repair` | Wrong column index, unanchored drifting range, missing exact-match flag |
| `budget-variance-audit` | Sign-convention inconsistency, hardcoded total, wrong percentage base |
| `headcount-restraint` | Restraint: one broken cell in a healthy column, where the cheap failure is rewriting the column |

`headcount-restraint` deserves a note. Its target is a single cell, so `Modif.`
is easy. It exists to measure whether an agent leaves working formulas alone,
which is the property that separates a usable spreadsheet agent from a
dangerous one.

## Current limits

- **No recalculation oracle.** Grading is formula-level and value-level against
  declared expectations. Seeds intentionally omit cached values, and the
  validator now rejects value assertions on formula cells rather than letting
  them fail mysteriously. Verifying that a formula *computes* the right answer
  needs a headless recalc engine (LibreOffice or a real Excel host).
- **No live host driver.** Results are graded from `.xlsx` files. Driving a
  real Excel session end-to-end is the next step, and is what would make these
  tasks true integration coverage rather than grader coverage.
- **Five tasks, all Excel, all synthetic.** Enough to gate against destructive
  behavior; not enough to characterize agent quality. Task count and realism
  are the things to grow, and no synthetic seed reproduces the scale or
  weirdness of a real client workbook.
- **No per-variant pass-rate tracking over time.** The harness grades a run; it
  does not yet store a history to detect regression across model versions.
- **Formula equality is textual.** Two algebraically equivalent formulas are
  graded as different. That is deliberate for restraint tasks, but it means the
  grader cannot credit a legitimate refactor.
