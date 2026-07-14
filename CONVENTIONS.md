# Conventions & Workflow — MEWC Lambda Library

The single source of truth for how the lambdas and VBA in this repo are written
and how the repo syncs with Excel. If you use an AI coding assistant, point it at
this file.

## Repo layout
- `lambdas/<name>.lambda` — one file per lambda; the canonical definition.
- `vba/*.bas` — VBA modules (imported into Excel via the VBE).
- `tools/lambda_check.py` — authoring-rule checker; run before committing.
- `MEWC Lambda and VBA Unit Tests.xlsm` — the committed test workbook (formatted Prep + all VBA + lambdas + unit-test sheets). The upstream/dev workbook, refreshed from the text sources — the text files are what you edit. The downstream template lives outside the repo.

### `.lambda` file format
Four delimited sections; `SIGNATURE`, `COMMENT`, `CODE` are required. A blank line
precedes every header except the first (`SIGNATURE`).
```
=== SIGNATURE ===
name(arg1, [opt2])

=== COMMENT ===
<= 255 char one-liner (feeds Excel's Name comment)

=== CODE ===
=LAMBDA(...)

=== DESCRIPTION ===
<full description>
```
The lambda's name is the text before `(` in the signature and must match the
filename. The parsers (`lambda_check.py`, `import_lambdas`) tolerate the blank
lines either way; `export_lambdas` writes them so in-Excel edits keep the style. **The `DESCRIPTION` section is the documentation** — there is no
separate descriptions file to keep in sync.

## The hard rules about writing workbooks

**1. Never let any tool write the competition template.** The template
(`MEWC Lambdas and VBA.xlsm`) lives **outside** the repo and only ever *reads*
finished lambdas from a local clone via VBA. No Python, no openpyxl, no automation
touches it — ever.

**2. Never write *any* `.xlsm` with openpyxl's `save()` (a full-file rewrite).**
A whole-workbook openpyxl save silently corrupts this workbook: it deletes the Prep
sheet's form-control buttons (`xl/drawings/drawing1.xml`), drops the rich values
(`xl/richData/*`) and dynamic-array metadata (`xl/metadata.xml` → implicit-intersection
`@` signs), reorders `styles.xml`/`sharedStrings.xml`, and can leave `#NAME?` on
`_xleta.*` names or orphaned phantom code-modules. openpyxl is fine for **reading**
(inspecting cells, the checker) — never for saving.

### Editing test cases in the test workbook — use `tools/xlsm_edit.py`
Test *cases* live only inside `MEWC Lambda and VBA Unit Tests.xlsm` (no text form yet),
so an AI tool that adds/edits them must write that binary. Do it **only** through
`tools/xlsm_edit.py`, which does **surgical string edits** on the individual `<c>` cells
you name and copies every other part byte-for-byte. That keeps Prep's buttons, the rich
values, all styles, and the VBA project intact. The only residue: cells you edit lose
their `cm=` dynamic-array marker (the `@` problem), which the workbook's own
`fix_test_formulas` / `lambda_update` VBA restores when you next open it in Excel. Cells
you don't touch keep their markers. Lambda **code** still flows text-first via
`import_lambdas` (see below) — you don't hand-edit the Lamb sheet for that.

## Workflow (repo ↔ Excel)
The committed `MEWC Lambda and VBA Unit Tests.xlsm` is the upstream/dev workbook; the text
files are what you **edit**. Keep the direction of truth straight — *edit text, refresh the
workbook, commit both*:

- **Lambdas** — edit `lambdas/*.lambda`, run `python tools/lambda_check.py
  lambdas/*.lambda`, then in the test workbook run `import_lambdas` (rebuilds the Lamb sheet
  + Name Manager). Commit the `.lambda` **and** the refreshed workbook.
- **Test cases** — edit the test sheets in `MEWC Lambda and VBA Unit Tests.xlsm` with
  `tools/xlsm_edit.py` (surgical cell edits; never openpyxl `save()`). Then open the
  workbook in Excel and run `lambda_update` / `fix_test_formulas` to restore dynamic
  arrays on the edited cells, confirm the tests pass, and save from Excel. Commit the
  workbook.
- **VBA** — edit `vba/*.bas`, then re-import the changed module into the test workbook
  (in the VBE: remove the old module, drag the `.bas` in — import won't overwrite). Commit both.
- `export_lambdas` writes the Lamb sheet back to `.lambda` files if you edited a
  lambda directly in Excel.

The downstream template workbook lives **outside** the repo and pulls finished lambdas from
a local clone — it's a consumer, never a source. See `MEWC Lambdas Edit and Test Workflow.md`
for the full edit → test → review → merge loop.

Never treat a workbook's baked lambdas/VBA as the source — always edit the text and re-import.

## Personal settings
Personal config (`Directory`, case/model filenames) is **not** in the repo — set it
per-machine at runtime.

## Address convention
Addresses are packed as one number: **numeric address = `1e6*row + col`**.
- `num_ad(x)` coerces an A1 string **or** a number to the numeric address
  (idempotent); `str_ad(x)` is the inverse.
- Any argument that takes a cell address should accept **either** form — coerce
  it with `num_ad` at the top and name it `*_ad`.
- Address-*producing* lambdas generally return numeric addresses.

---

# Authoring rules

Conventions for writing custom LAMBDAs so they paste into the Name Manager
cleanly and don't break. When generating a lambda, follow all of these.

## 1. One version only, no inline comments
- Produce a **single** copy of each lambda — the one that goes in the Name
  Manager. Do **not** keep a second "commented" copy (editing two drifts).
- **No comments of any kind inside the formula** — not `//`, not `#`, and not
  the `N("…")` trick. Excel formulas have no comment syntax; anything like that
  either errors or adds noise.
- Documentation lives **separately, in plain English** — in the `DESCRIPTION`
  section of the lambda's `.lambda` file (and the Comment/Description columns of
  the `Lamb` sheet). Each entry: name, purpose, arguments, and a short example.

## 2. Whitespace style — medium
Readable but compact:
- Start a **new line for each new variable** you define in a `LET`.
- **Do not** add a newline after every parenthesis.
- Indent nested `LET`/`LAMBDA` bodies one level.

```
=LAMBDA(arr,[as_row],
  LET(
    flat, TOCOL(arr, 3),
    clean, FILTER(flat, flat<>""),
    result, SORT(UNIQUE(clean)),
    IF(IF(ISOMITTED(as_row), 0, as_row),
       TOROW(result),
       result)
  )
)
```

## 3. Variable naming
Use **snake_case, no leading underscores** (`row_count`, `start_ref`, `dir`).

**Hard rules (these break the lambda if violated):**
- **Never use a name that is a valid cell reference.** The breaking pattern is
  1-3 letters immediately followed by digits, e.g. `x2`, `c1`, `num1`, `fy3`
  (`num1` is literally column NUM row 1). Treat any `letters+digits` token (≤3
  letters, no separator) as dangerous. (A token whose digits would be *row 0*,
  like `vec0`/`dx0`, is technically safe — row 0 isn't a real cell — but avoid
  the pattern anyway.)
  - No trailing digits. To distinguish two of something use words or an
    underscore-separated suffix: `val_a` / `val_b`, `first` / `second`, `ref_1`.
- **Don't reuse a built-in function name** (`N`, `T`, `C`, `R`, `SUM`, `ROW`,
  `LEFT`, `BASE`, `CHAR`…). Reusing one silently overwrites the built-in.
- **Don't reuse an existing defined name / lambda name** (`set`, `dist`,
  `arrow`, `rev`…). A LET variable that shadows a lambda changes behavior.

**Style rules (won't break, but keep it clean):**
- Single letters (`d`, `t`, `x`) are legal but avoid them; prefer short words.
- Short, digit-free words are ideal: `ref`, `dir`, `arr`, `row`, `col`, `flat`.
- Reuse the project's standard argument names so signatures stay consistent.
- *Legacy note:* several older lambdas (e.g. `route_cost`, `z_route_worker`,
  `freq_table`) use leading-underscore names. They're grandfathered — leave them;
  write **new** lambdas without leading underscores.

## 4. Optional arguments — square brackets
Optional parameters **must be wrapped in square brackets** in the parameter list;
forgetting the brackets makes the argument required.
```
=LAMBDA(arr, [as_row], …)      correct
=LAMBDA(arr, as_row, …)        wrong — as_row is now required
```
Test optional args inside with `ISOMITTED`, and put required args first, all
optional (bracketed) args after them.

## 5. Stay under 2048 characters
The Name Manager caps a definition at **2048 characters**, and whitespace counts.
If a lambda approaches the limit:
- **Reuse** an existing library lambda where possible (this is how `set_u` stays
  tiny by calling `set`), then **decompose** a reusable inner piece into its own
  named lambda, then trim indentation. Prefer decomposition over dense minifying.
- If a formula is genuinely close to the cap, note the character count in its
  description so future edits stay aware.

## 6. Private helper lambdas — `z_` prefix
Helpers that exist only to support another lambda (e.g. the recursive workers
behind `xbyrow` and `route_cost`) get a **`z_` prefix** (`z_xbyrow_ix`,
`z_route_worker`).
- **Sorts to the bottom.** `z_` pools helpers at the end of the (alphabetical)
  Name Manager and AutoComplete. A leading `_` sorts to the *top*; avoid it.
- **Keeps autocomplete clean.** Typing the public name never surfaces the helper.
- **Call only the public wrapper** from worksheets; treat any `z_` name as
  internal.

## 7. Editing lambdas — keep Code cells as text
Lambdas live in `.lambda` files (canonical) and, in Excel, the `Lamb` sheet
(Signature, Comment, Code, Description); `lambda_update` pushes them into the
Name Manager.
- Each `Lamb` **Code** cell is stored as **text with a leading apostrophe** (quote
  prefix) so `=LAMBDA(...)` stays text and doesn't flip into a live formula. The
  apostrophe is invisible in the cell but shows in the formula bar.
- If you edit a Code cell directly, **keep the leading apostrophe**; retyping
  `=LAMBDA(...)` without it converts it to a formula (`@`, `#NAME?`).
  (`import_lambdas` handles this for you when syncing from the repo.)

## 8. Error handling — catch `#N/A`, keep real errors visible
If a lambda's output can produce **`#N/A`**, wrap the public output in
`IFNA(x, "")` so a stray `#N/A` becomes a blank instead of poisoning everything.
- **Why `#N/A` specifically.** It's usually a benign edge case — `VSTACK`/`EXPAND`
  padding, an `XLOOKUP`/`XMATCH` miss, an empty `FILTER`. `IFNA` replaces those
  **element-wise**, so the good cells survive.
- **Use `IFNA`, not `IFERROR`.** `IFNA` leaves `#NAME?`, `#REF!`, `#SPILL!`,
  `#VALUE!` visible — those usually mean something is actually broken.
- **Wrap the outermost/public output, and add it last.** `z_` helpers can stay
  pure and let the public wrapper clean once.
- **Remember `""` is text** — fine for display/joins, but arithmetic on a blanked
  cell re-errors, so this suits assembly/display outputs more than numeric ones.

## 9. Return a true scalar, not a 1×1 array
When a result is a single value, make sure it's a genuine **scalar**, not a
**1×1 array** — they look identical in a cell but differ when fed to an
array-shaping function.
- **The trap.** `SEQUENCE(only_numbers("X10"))` returned `1`, not `1..10`,
  because `only_numbers` handed back `{10}` (1×1) not the scalar `10`; `SEQUENCE`
  runs per element and keeps only each top-left value.
- **Where they sneak in.** `REGEXEXTRACT` (return-mode 1), `FILTER`, `TEXTSPLIT`,
  `UNIQUE` stay "array" for one element; an outer `IF`/`MAP` whose *condition* is
  an array re-wraps a scalar into a 1×1 array.
- **The fix.** Unbox with `INDEX(x, 1)` (or `@x`), and guard the "single vs many"
  split with a **scalar** condition like `ROWS(x) * COLUMNS(x) = 1`.

## 10. Pre-commit checklist
1. No inline comments anywhere.
2. Every variable is snake_case, no leading underscore (new lambdas), **no
   letter+digit tokens**, and collides with no function or defined name.
3. One line per `LET` variable; no newline-after-every-paren noise.
4. All optional parameters are in `[brackets]` and handled with `ISOMITTED`.
5. Definition length is under 2048 characters (decompose if not).
6. A plain-English `DESCRIPTION` exists in the `.lambda` file, not in the formula.
7. Private helpers (called only by another lambda) use the `z_` prefix.
8. If the output can return `#N/A`, it's wrapped in `IFNA(x, "")` (not `IFERROR`).
9. A single-value result is a true scalar (`INDEX(x, 1)`/`@x`), not a 1×1 array.
10. `python tools/lambda_check.py lambdas/<name>.lambda` passes.
