# MEWC Lambdas Edit and Test Workflow

## Overview

When a lambda is created or edited, complete testing in Excel and review the code
diff **before** merging into the remote repo. The two workbooks play different roles:

- **`MEWC Lambda and VBA Unit Tests.xlsm`** — the *upstream / dev* workbook. Lives in the
  repo; holds the test worksheets and the `unit_test_tools` macros. This is where you edit and test.
- **The template workbook** (`MEWC Lambdas and VBA.xlsm`) — the *downstream* consumer.
  Lives **outside** the repo and pulls finished lambdas from a local clone. No test sheets.

The **source of truth is the `lambdas/*.lambda` text files**, not either workbook. The
Lamb sheet and the `.lambda` files are two views of the same thing and must be kept in
sync; `import_lambdas` / `export_lambdas` do that syncing.

---

## Editing (by user or AI tool)

Pick **one** lane per edit and stick to it — never hand-edit the Lamb sheet *and* the
`.lambda` file for the same change, or the two will drift and the next sync will clobber
one of them.

**Lane A — text-first (use when an AI tool such as Claude Code makes the edit):**

1. Edit the signature / code / comment / description in `mewc-lambdas\lambdas\<name>.lambda`.
2. In `MEWC Lambda and VBA Unit Tests.xlsm`, run `import_lambdas` — it rewrites the Lamb
   sheet from the `.lambda` files and pushes the names into the Name Manager.

**Lane B — Excel-first (use when you hand-edit in the workbook):**

1. Edit the signature / code / comment / description on the **Lamb** tab.
2. Run `export_lambdas` — it writes the Lamb rows back out to `lambdas\<name>.lambda`.

Either lane leaves the Lamb sheet and the `.lambda` file matching. Then:

3. Add or edit a testing worksheet in `MEWC Lambda and VBA Unit Tests.xlsm` with enough
   test cases to cover the new or changed behavior.

---

## Testing (by user)

1. **Reconcile first.** Make sure the Lamb sheet and the `.lambda` files match before you
   test — you want to test exactly what you'll commit. If you edited in Excel (Lane B),
   run `export_lambdas` now. (Lane A already reconciled them via `import_lambdas`.)
2. Run `lambda_update`. It copies the lambda code from the Lamb sheet into the Name
   Manager, hides the `z_` helpers, and rewrites the test-sheet formulas as dynamic arrays
   (stripping any stray `@` left over from AI editing).
3. Confirm the tests pass **and** that they actually exercise the new/changed
   functionality. The `test_summary` sheet gives the per-lambda pass/fail roll-up.

> Excel testing is a **manual** gate — it can't run in CI (there's no Excel on GitHub
> Actions). CI only re-checks the text authoring rules; a lambda that passes the checker
> but computes the wrong answer will only be caught here.

---

## Review Code Diff and Pull into Repo

"Pull into the repo" here means **push your branch and merge the PR** — the diff review
happens in the pull request before anything lands on `main`.

**Prerequisite:** the `.lambda` files reflect the tested code. If you edited in Excel
(Lane B), you already ran `export_lambdas` in the Testing step, so they do.

```bash
cd mewc-lambdas
git checkout -b edit/<name>                        # e.g. edit/xrank-tocol

python tools/lambda_check.py lambdas/*.lambda      # local authoring-rule gate — must pass
git add -A
git commit -m "<name>: <what changed>"
git push -u origin edit/<name>

gh pr create --fill                                # CI (lambda-check) runs on the PR
# Review the diff in the PR. Note in the PR description that the Excel tests passed.
gh pr merge --squash --delete-branch
```

Claude Code can drive most of this: make the `.lambda` edit, run the checker, walk you
through `git diff`, draft the commit message, open the PR with `gh`, and `/review` the
diff. The two steps it can't do stay with you: running `import_lambdas` and eyeballing the
tests in Excel.

---

## Update Template File

The template lives **outside** the repo and pulls from its own local clone, so update the
clone first:

1. In the template's local clone of `mewc-lambdas`: `git pull`.
2. In the template workbook, run `sync_template_from_repo` to load the updated lambdas from
   the clone. *(Sub still to be written — the downstream equivalent of `import_lambdas`,
   without the test-formula fixing.)*

---

## What the repo tracks

`.gitignore` tracks the **test workbook** (`MEWC Lambda and VBA Unit Tests.xlsm`) and ignores
every other `*.xlsm`, including the template, which lives outside the repo. The test *cases*
exist only inside that binary — there's no text form — so committing the workbook is how they
get versioned. The tradeoff is an opaque binary in the history (no readable diff on the
workbook itself); the meaningful diff is always the `.lambda` text. A future text export of
test cases would give diffable tests, but that's a separate project.
