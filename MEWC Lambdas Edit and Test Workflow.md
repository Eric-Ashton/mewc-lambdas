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
sync; `sync_test_workbook_from_repo` (repo → workbook) and `repo_export.export_lambdas`
(workbook → repo) do that syncing.

---

## Editing (by user or AI tool)

Pick **one** lane per edit and stick to it — never hand-edit the Lamb sheet *and* the
`.lambda` file for the same change, or the two will drift and the next sync will clobber
one of them.

**Lane A — text-first (use when an AI tool such as Claude Code makes the edit):**

1. Edit the signature / code / comment / description in `mewc-lambdas\lambdas\<name>.lambda`.
2. In `MEWC Lambda and VBA Unit Tests.xlsm`, run `sync_test_workbook_from_repo.sync_lambdas`
   — it rewrites the Lamb sheet from the `.lambda` files, pushes the names into the Name
   Manager, and strips any stray `@`. (Run `sync_all` instead to also re-import the VBA and
   run the tests.)

**Lane B — Excel-first (use when you hand-edit in the workbook):**

1. Edit the signature / code / comment / description on the **Lamb** tab.
2. Run `export_lambdas` — it writes the Lamb rows back out to `lambdas\<name>.lambda`.

Either lane leaves the Lamb sheet and the `.lambda` file matching. Then:

3. Add or edit test cases on the lambda's worksheet in `MEWC Lambda and VBA Unit
   Tests.xlsm` to cover the new or changed behavior. Claude Code can do this directly with
   `tools/xlsm_edit.py`, which surgically rewrites only the cells it changes and leaves
   every other part of the workbook (Prep's buttons, rich values, styles, VBA) byte-for-byte
   intact — the one thing a plain openpyxl save would wreck. The edited cells lose their
   dynamic-array marker (the `@` problem); the Excel-side `lambda_update` /
   `fix_test_formulas` step below restores them. You can also add test cases by hand in Excel.

---

## Testing (by user)

1. **Reconcile first.** Make sure the Lamb sheet and the `.lambda` files match before you
   test — you want to test exactly what you'll commit. If you edited in Excel (Lane B),
   run `repo_export.export_lambdas` now. (Lane A already reconciled them via `sync_lambdas`.)
2. Run `lambda_update`. It copies the lambda code from the Lamb sheet into the Name
   Manager and rewrites the test-sheet formulas as dynamic arrays
   (stripping any stray `@` left over from AI editing).
3. Confirm the tests pass **and** that they actually exercise the new/changed
   functionality. The `lambda_tests` sheet gives the per-lambda pass/fail roll-up.

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

Claude Code can drive most of this: make the `.lambda` edit, author/edit the test cases in
the workbook via `tools/xlsm_edit.py`, run the checker, walk you through `git diff`, draft
the commit message, open the PR with `gh`, and `/review` the diff. What stays with you is
the Excel side: running `sync_test_workbook_from_repo.sync_all` (Lamb + Name Manager + `@`-fix
+ tests; there's no Excel on GitHub Actions), then eyeballing that the tests actually pass before you commit
the workbook.

---

## Update Template File

The template lives **outside** the repo and pulls from its own local clone, so update the
clone first:

1. In the template's local clone of `mewc-lambdas`: `git pull`.
2. In the template workbook, run `sync_template_from_repo.sync_all` (from the VBE — it's an
   `Option Private Module`, so it's hidden from Alt+F8). It loads the updated lambdas into the
   Lamb sheet + Name Manager and re-imports every shared/template VBA module (pruning any
   stale one), skipping the test-only tooling and the test-formula fixing.

---

## What the repo tracks

`.gitignore` tracks the **test workbook** (`MEWC Lambda and VBA Unit Tests.xlsm`) and ignores
every other `*.xlsm`, including the template, which lives outside the repo. The test *cases*
exist only inside that binary — there's no text form — so committing the workbook is how they
get versioned. The tradeoff is an opaque binary in the history (no readable diff on the
workbook itself); the meaningful diff is always the `.lambda` text. A future text export of
test cases would give diffable tests, but that's a separate project.
