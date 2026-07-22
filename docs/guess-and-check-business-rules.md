# Guess-and-Check — Business Rules

Reverse-engineered from `guess and check.xlsm` (the standalone strategy workbook)
so the same behaviour can be rebuilt as a button on the competition template's
`_LX` level sheets. This document is the authoritative spec: it describes **what
the workbook does**, cell by cell and rule by rule, not how the current VBA
happens to be written.

---

## 1. Purpose and when to use it

In the MEWC live web platform (`play.excel-esports.com`) the scoreboard reports
**points per level almost instantly** after you submit. A level is a set of games
(e.g. Level 3 = games 31–50), each worth the same fixed number of points, and the
platform tells you your **total points for that level** — i.e. how many of your
current answers are correct — but **not which ones**.

When a level's answers are **integers with few significant digits**, it can be
faster to *guess and check* against that point feedback than to actually solve the
case. This tool automates that search: you enter a whole column of guesses (one per
game), submit them on the platform, read back how many were right, click one of
three feedback buttons, and the tool produces the next column of guesses. Repeat
until every game is solved.

> This is a fallback/opportunistic strategy. Genuinely solving the case is still
> preferred; guess-and-check only pays off when answers are small integers and the
> platform gives fast per-level feedback.

---

## 2. The information model (why it works)

- One **level** contains `G` games; each correct answer is worth `P` points, so a
  level is worth `G × P`.
- After a submission the platform reveals **`points`**, hence
  **`correct = points / P`** — the *count* of correct guesses in the batch, with no
  attribution to specific games.
- The tool therefore runs **two nested searches at once**, both driven only by that
  count:
  1. **Value search (per game).** Every unsolved game holds one candidate value in
     the *Guesses* column. Each round eliminates the values that were just proven
     wrong and steps outward, so the candidate for each game marches away from a
     central starting guess in units of the answer's significance.
  2. **Attribution search (which game).** When a batch scores `1` or `2+` correct,
     you know *how many* guesses in the column were right but not *which*. The tool
     bisects the column — holding half of the guesses aside in a "Possible" bucket
     and re-testing the other half — to isolate the winning games over several
     rounds.

The three feedback buttons correspond to the only three cases that matter for the
next step: **0 correct**, **exactly 1 correct**, **2 or more correct**.

---

## 3. Workbook structure

| Sheet | Role |
|---|---|
| `case data` | A verbatim paste of the competition case (narrative, bonus questions, and a game→level→points→answer table). Read-only input; the level sheets look everything up from here. |
| `L1` … `L7` | One working sheet per level. Identical layout; each is pinned to its level number and self-configures from `case data`. |

Rebuild target: instead of seven pre-made sheets, a button on each template `_LX`
sheet should **generate** a working sheet (or region) with the same layout and
formulas, pre-pointed at that level.

---

## 4. The `case data` sheet (input)

A pasted copy of the case. The setup formulas depend on a table whose columns are:

| Column | Meaning |
|---|---|
| **B** | Game label — a game number, an `ExampleN:` marker, or `Bonus N`. |
| **C** | The level the row relates to (numeric level number for real games). |
| **D** | Points per game for that level. |
| **E** | The answer (used for the per-level worked *example*). |

Only B/C/D/E are consumed. Everything else on the sheet is reference text for the
human solver.

---

## 5. Setup block (rows 2–11 of an `L` sheet)

`A` holds the label; `B` (and `C` for hint ranges) holds the value/formula. All
lookups target the `case data` sheet.

| Cell | Label | Value / formula | Meaning |
|---|---|---|---|
In the automated flow (§11) the button writes every value below by **parsing the
source `_LX` sheet**; the operator can then override any of them. The formulas
shown are the *reference* logic — the button may hard-write parsed values instead
of live formulas, but the meaning is identical.

| `B2` | Level N | the sheet's level number (1–7) | Selects which level this sheet works. Everything else derives from it. Parsed from the `_LX` sheet name / header (§11). |
| `B3` | Points Per Game | `=XLOOKUP(B2, 'case data'!C:C, 'case data'!D:D, , , -1)` → `P` | Points each correct answer is worth. Parsed from the case table's level section (§11). |
| `B4` / `C4` | Game Ns Range | `=LET(qs, FILTER('case data'!B:B, 'case data'!C:C=B2), MIN(qs))` and `MAX(qs)` | First and last game number in the level. Count `G = C4 - B4 + 1`. |
| `B5` | Example Answer | `=XLOOKUP("Example"&B2, …)` (see §11.2 for multiple examples) | The worked-example answer(s) for the level — seeds the guess center and significance. A level may expose **more than one** example (`Example5a`, `Example5b`); use the **mean** of them. Must be numeric — see the pre-flight guard (§11.3). |
| `B6` | Level Significance | see below → `sig` | The step size / smallest significant unit of the answers (e.g. `1` for integers, `0.01` for cents). **Inferred, user-overridable** (§11.5). |
| `B7:C9` | Hint 1/2/3 Range | (min in `B`, max in `C`) | Up to three `[min,max]` ranges narrowing where answers lie. **Parsed from the hint text** (§11.4); blank when a level has no usable range hints. |
| `B10` | Guess Center | 50/50 of example(s) and hints — see §11.6 | The starting guess for games that have no hint of their own. **Half the mean of the example answer(s) plus half the mean of the usable hint boundaries**, rounded with `CEILING.MATH(…, sig)`; with no usable hints it is just the mean of the example(s). |
| `B11` | Negative Allowed | `0`/`1` | Whether answers may be negative (bounds the downward search). **Inferred, user-overridable** (§11). |

**Significance formula (`B6`)** — derive the smallest power of ten by which the
example answer is exactly divisible, i.e. its significance:

```
=LET(
  x, B5, x_abs, ABS(x),
  power_list, SEQUENCE(21, , -10, 1),      ' powers -10..10
  sig_list, 10 ^ power_list,               ' 1e-10 .. 1e10
  ceil_list, IFERROR(CEILING.MATH(x_abs, sig_list), NA()),
  matches, IF(ceil_list = x_abs, sig_list, ""),
  MAX(matches))                            ' largest step that divides x exactly
```

So `B5 = 20 → sig = 1`; `B5 = 439 → sig = 1`; `B5 = 0.05 → sig = 0.01`.

**Row 12** is a diagnostics row: each of `A12:H12` is `=COUNT(<col>14:<col>1000)`,
i.e. the live count of numbers in that column (games, correct answers, guesses,
etc.).

---

## 6. Feedback buttons (labels)

Three cells render human-readable button captions from `P` and the game count
`G = (C4 - B4) + 1`:

| Cell | Formula | Renders (e.g. P=7, G=20) | Bound macro |
|---|---|---|---|
| `E8` | `="0 / " & B3*((C4-B4)+1) & " Points"` | `0 / 140 Points` | `zero_right` |
| `F8` | `=B3 & " / " & B3*((C4-B4)+1) & " Points"` | `7 / 140 Points` | `one_right` |
| `G8` | `=2*B3 & "+ / " & B3*((C4-B4)+1) & " Points"` | `14+ / 140 Points` | `two_or_more_right` |

The captions tell the operator which button matches the score the platform just
showed (0 points → 0 right, `P` points → 1 right, `≥2P` points → 2+ right).

---

## 7. Working table (header at "Game Numbers" row; data below)

Columns are found by role, not by fixed row — the code locates the header row by
searching column A for the text **"Game Numbers"**, then works from the next row
down to the last populated game.

| Col | Header | Written by | Meaning |
|---|---|---|---|
| **A** | Game Numbers | formula | `=SEQUENCE(C4-B4+1, 1, B4)` — one row per game, numbered `B4…C4`. |
| **B** | Correct Answers | macros | The confirmed answer for a solved game. **Blank = unsolved.** A numeric value here (including `0`) means solved. |
| **C** | Guesses | macros | The candidate value to enter on the platform **this** round, per game. The whole column is what you submit. |
| **D** | Eliminated Min | macros | Low edge of the contiguous block of values already proven wrong for this game. |
| **E** | Eliminated Max | macros | High edge of that eliminated block. The answer lies **outside** `[D,E]`; new guesses probe just past the edges (`D - sig` or `E + sig`). |
| **F** | Possible A (≤ 1 Right) | macros | "Held-out" candidate bucket used by the attribution search (see §8). |
| **G** | Possible B | macros | Second held-out bucket, used when 2+ were right. |
| **H** | Initial Guesses | formula | The seed guess for each game before any elimination. |

**Initial-guess formulas (`H`)** — the first three data rows key off the three hint
ranges; the rest use the guess center:

```
H(row1) = IF(B7="", B10, CEILING.MATH(AVERAGE(B7:C7), B6))
H(row2) = IF(B8="", B10, CEILING.MATH(AVERAGE(B8:C8), B6))
H(row3) = IF(B9="", B10, CEILING.MATH(AVERAGE(B9:C9), B6))
H(rowN) = IF(A(rowN)="", "", B10)          ' N >= 4
```

`B`, `C`, `D`, `E`, `F`, `G` carry **no formulas** — they are pure state written and
cleared by the three macros.

---

## 8. Feedback macros — exact rules

Shared preamble for all three:

- `ws = ActiveSheet`; read `sig = B6`.
- Locate header row by finding `"Game Numbers"` in column A; `firstRow = header+1`,
  `lastRow =` last non-empty cell in column A.
- Column map: A=Game, B=Correct, C=Guess, D=ElimMin, E=ElimMax, F=PossA, G=PossB,
  H=InitGuess.
- **`0` is always a valid answer** — "solved" means *numeric and non-empty*, never
  "non-zero".
- After building the next `Guesses` column, the macro **copies** the guess range to
  the clipboard so the operator can paste it straight into the platform.
- A "solved" game (numeric `B`) is skipped when generating new guesses; its guess
  stays blank.

### 8.1 `zero_right` — batch scored 0 correct

Every guess in the column was wrong. Record that, then decide what to test next.

1. **Snapshot** the starting counts *before* touching anything:
   `cntA0 =` number of non-empty `Possible A` cells; `cntB0 =` number of non-empty
   `Possible B` cells. Remember the row of the first `Possible A` (`idxSingleA0`).
2. **Process each guess** (all now known wrong):
   - If the game has no eliminated block yet (`D` and `E` both empty), set
     `D = E = guess` (the guess becomes the eliminated block).
   - Otherwise extend the block **only if the guess sits exactly one step past an
     edge**: if `guess = D - sig` then `D = guess`; if `guess = E + sig` then
     `E = guess`.
   - If the guess equals the game's `Possible A` or `Possible B` value, clear that
     bucket cell (that held-out candidate is disproven).
   - Clear the guess cell (the column is rebuilt below).
3. **Branch on the snapshot:**
   - **Case A — `cntA0 = 1` and `cntB0 = 0`:** the single surviving `Possible A`
     value is the answer for its game → copy it into `Correct Answers`, clear that
     `Possible A`. Then **generate fresh guesses** for every still-unsolved game
     (see §8.4). Done.
   - **Case B — `cntA0 ≥ 2` or `cntB0 ≥ 1`:** you are mid-bisection. Move **half
     (rounded up)** of the `Possible A` values back into `Guesses`, and if
     `cntB0 > 0` also move half (rounded up) of `Possible B` into `Guesses`.
     Clear those buckets' moved cells. **Generate no new guesses** — you are only
     re-testing held-out candidates. Done.
   - **Case C — `cntA0 = 0` and `cntB0 = 0`:** ordinary elimination round →
     **generate fresh guesses** for every unsolved game (see §8.4). Done.

### 8.2 `one_right` — batch scored exactly 1 correct

Exactly one guess in the column is right; find which by bisection.

1. `Possible B` is irrelevant here → clear the whole `Possible B` column.
2. Count the guesses; remember the first guess row.
3. **If exactly one guess was in the column:** that guess *is* the answer → copy it
   into `Correct Answers`, clear all `Guesses` and all `Possible A`, then
   **generate fresh guesses** for every unsolved game (§8.4). Done.
4. **If 2+ guesses were in the column** (the normal bisection step):
   - First, fold any existing `Possible A` into the eliminated ranges using the same
     "one step past the edge" rule as in `zero_right` step 2 (if `possA = D - sig`
     then `D = possA`; if `possA = E + sig` then `E = possA`).
   - Clear the whole `Possible A` column.
   - Move **half (rounded down)** of the current guesses into `Possible A` (holding
     them out), leaving the rest in `Guesses` to re-test. **No new guesses.** Done.

### 8.3 `two_or_more_right` — batch scored 2+ correct

At least two guesses are right; hold half aside and keep testing.

1. Clear the whole `Possible A` column (unused in this path).
2. Move **half (rounded down)** of the current guesses into `Possible B`, leaving the
   rest in `Guesses`. **Never generate new guesses.** Done.

### 8.4 "Generate fresh guesses" subroutine

For each game from `firstRow` to `lastRow`:

- If `Correct Answers` is numeric → solved, leave the guess blank.
- Else if the game has **no** eliminated block yet (`D` empty) → guess `= Initial
  Guesses (H)`.
- Else → step just outside the eliminated block, choosing an edge **at random**:
  50% `guess = D - sig`, 50% `guess = E + sig`.

(`Randomize` is called once per macro; the random edge choice keeps the search from
biasing consistently high or low.)

---

## 9. Operator workflow (the loop)

1. Open the `L` sheet for the level; confirm `B2` is the level number and the setup
   block populated (games, points, center, significance) from `case data`. Fill in
   any known `Hint 1/2/3` ranges (`B7:C9`) if you have them.
2. The `Guesses` column starts at the initial guesses. **Copy it and paste into the
   platform's answer boxes for that level; submit.**
3. Read the level's points off the scoreboard → divide by `P` to get how many were
   correct.
4. Click the matching button: `0 / … Points` (`zero_right`), `P / … Points`
   (`one_right`), or `≥2P / … Points` (`two_or_more_right`). The macro updates state
   and puts the **next** guess column on the clipboard.
5. Paste, submit, repeat. Games move into `Correct Answers` as they're pinned down;
   when every game is solved the `Guesses` column is empty and `Correct Answers` is
   full.

---

## 10. Conventions and edge cases (must preserve on rebuild)

- **`0` is a valid answer.** Every "is this solved / is this a real value" test must
  be *numeric-and-present*, not *non-zero*. This is the single most common way a
  reimplementation breaks.
- **Significance `sig` is the step unit.** All probing moves by `± sig`, and initial
  guesses are rounded with `CEILING.MATH(…, sig)`. Non-integer answers work if `sig`
  is set correctly (e.g. `0.01`).
- **`[ElimMin, ElimMax]` is the eliminated block**, not the feasible range. The
  answer is strictly outside it; guesses probe the two values immediately outside.
  Extension only happens when a guess/possible is exactly one `sig` past an edge —
  eliminations never "jump".
- **Possible A vs Possible B** are two independent hold-out buckets for the
  attribution bisection. `one_right` only ever uses A; `two_or_more_right` only ever
  uses B; `zero_right` can drain either when resuming a bisection (Case B).
- **Rounding of the "half" moved differs by macro on purpose:** `zero_right` Case B
  moves half **rounded up**; `one_right` and `two_or_more_right` move half **rounded
  down**. Preserve each exactly — it controls convergence.
- **Header is found by text**, and the row range by the last populated game number,
  so the table can start on any row and be any length.
- **Negative Allowed (`B11`)** is intended to bound the downward search; the current
  macros read `sig`/center but do not yet clamp at zero — treat clamping as an
  intended rule to (re)confirm with the author before relying on it.

---

## 11. Automation — the "Guess-and-Check" button

The target implementation is a **button on each competition-template `_LX` level
sheet**. Pressing it builds a fresh working sheet named **`_GCN`** (Guess-and-Check
for level N — e.g. on `_L6` it makes `_GC6`) with the §5–§7 layout, then **parses
the level's data off `_LX` (and the imported case) and pre-fills the setup block**,
so the operator can start the loop immediately and only correct what the parse got
wrong.

### 11.1 Trigger and sheet creation

1. Determine the level number **N** from the active `_LX` sheet — either the sheet
   name (`_L6` → `6`) or its header cell (e.g. `B3 = "Level 6"`). Sheet name is the
   more reliable key.
2. Target sheet name **`_GCN`** (`"_GC" & N`). **Never overwrite an existing one —
   always create a fresh copy.** If `_GCN` is free, use it; otherwise use the lowest
   free suffixed name `_GCN(2)`, `_GCN(3)`, … (e.g. pressing the button on `_L6`
   twice gives `_GC6` then `_GC6(2)`). This preserves any in-progress sheet, so no
   confirm prompt is needed.
3. Create the sheet, lay down the setup block (rows 2–11), the feedback-button
   captions (`E8:G8`), and the working-table header + formula columns (`A` game
   numbers, `H` initial guesses). `B`–`G` start empty (macro state).

### 11.2 Where the data lives on an `_LX` sheet

Layout observed on the reference template (confirm against the live template — it
may have shifted):

| Datum | Location on `_LX` | Notes |
|---|---|---|
| Level number, difficulty | `B3` = `"Level 6"`, `C3` = `"Hard"` | Parse `N` from the trailing integer. |
| Worked-example answer(s) | the table row(s) whose `Game #` (col B) is `"ExampleN"` (or `ExampleNa`, `ExampleNb`, …); read the **Answer** (col E) of each | e.g. `E21 = 87`. A level may have **more than one** example row — collect them all. Prefer these numeric cells over the prose in `B9:B13`. |
| Hint lines | column **C** just above the table (e.g. `C15:C17`) | Free text, one per hinted game — see §11.4. |
| Game numbers | column **A** below the table (e.g. `A27… = 91,92,…`), or the case level section | The full list of games for the level. |

Per-game **points** are **not** cleanly on `_LX` (its example row shows `0`). Read
them from the imported case sheet (`Case` / `case copy` / equivalent), which has,
per level, a `Level N Header` row (`Game # | Level | Points | Answer`) followed by
`Example` (points `0`) and the real game rows. Take the **Points of any real game
row for level N** (they're all equal — level 6 = `9`); do **not** use the example
row's `0`. Equivalent lookup: points = the Points column value where Level = N and
`Game #` is numeric.

### 11.3 Pre-flight: numeric-answer guard (abort condition)

Guess-and-check only works when a level's answers are **single numbers**. Before
building anything, check the example answer(s):

- If every example answer is a single numeric value (integer or decimal, optionally
  negative) → proceed.
- If any example answer is **non-numeric** — text, alphanumeric (e.g. an A1
  reference), or a **delimited series of numbers** (e.g. `"20, 10, 20"`) — **abort
  immediately** with the message:

  > **Guess and Check only works on numeric answers**

  Do not create the `_GCN` sheet in this case.

(The example-answer *prose* in `B9:B13` often lists several numbers; that is why the
guard reads the clean **Answer cell** from the example table row, not the narrative.
`Example6` answer `87` is a single number → OK; a level whose answer is `"A5"` or
`"3;7;9"` → aborts.)

### 11.4 Parsing the hints

Each usable hint is a sentence of the form:

```
Game #91 Hint: The correct answer is between 90 and 100.
```

For every hint line found (in the `_LX` hint column, or the case's `Level N Instr`
rows):

- **Game number** = the integer after `"Game #"`.
- **min / max** = the two numbers in the `"between <min> and <max>"` phrase (first
  and second numeric tokens after `between`). Parse as numbers.
- **A hint is "usable" only if it matches that numeric-range pattern.** Hints that
  don't — e.g. `"the answer is odd"`, `"the answer is a multiple of 5"`, or any
  level-1-style prose with no `between x and y` — are **ignored** (treated as if
  that game had no hint). A level with zero usable hints is handled exactly like a
  level with no hints at all (see §11.6, special case).

Map each parsed `[min,max]` to `Hint 1/2/3 Range` (`B7:C9`) in the order the hints
appear, which is also the order of the level's first hinted games. (The reference
workbook assumes ≤ 3 usable hints covering the level's leading games; if a level
ever hints specific non-leading games, place each hint's range against its matching
game row instead.) A pre-parsed `Hints` sheet with columns `hint text | game# | min
| max | answer | sig` exists in the template and can be used as a cross-check or
shortcut.

### 11.5 Inferring significance (`B6`) — user-overridable

Compute `sig` = the largest power of ten (from `1e-10` to `1e10`) that divides the
sample values **exactly**, using the §5 significance formula. Feed it the example
answer(s) and, for robustness, the usable hint `min`/`max` values: take the
**coarsest common significance** across all of them (the largest step that divides
every sample). E.g. `87, 90, 100, 350 → sig = 1`; `0.05, 0.10 → sig = 0.01`. Write
the inferred value to `B6` but leave it editable — the operator confirms/overrides
before the first submission.

### 11.6 Initial guess / Guess Center (`B10`)

The starting guess for a game with **no hint of its own** is the **Guess Center**,
computed as a 50/50 blend:

```
ex  = mean of all example answers (§11.2)             ' >= 1 example, all numeric
hb  = mean of all usable hint boundary numbers        ' 2 per usable hint (min+max)

guess_center = ex                       if there are NO usable hints
             = 0.5 * ex  +  0.5 * hb     otherwise

B10 = CEILING.MATH(guess_center, sig)
```

So examples and hints get **equal total weight**, regardless of how many of each
there are. Special cases this covers:

- **No usable hints** (special case: none present, or all non-numeric per §11.4) →
  `guess_center = ex`, and the algorithm simply runs with no hint bounds.
- **Multiple examples** (`Example5a`, `Example5b`) → `ex` is their mean; still 50/50
  against the hint mean (or just `ex` if no usable hints).

A game that **does** have its own parsed hint uses that hint's own midpoint as its
initial guess instead of the center: `CEILING.MATH(AVERAGE(hint_min, hint_max),
sig)`. Non-hinted games fall back to `B10`. (This is the `H`-column logic in §7.)

### 11.7 Inferring "Negative Allowed" (`B11`) — user-overridable

Scan the example answer(s) and every usable hint `min`/`max`. **If none of them is
negative, set `Negative Allowed = 0` (No).** If any sample is negative, set it to
`1`. Write it to `B11` and leave it editable so the operator can flip it when they
know a level's answers can go negative even though the visible samples don't.

### 11.8 After the build

`_GCN` (or `_GCN(k)`) is now a standalone guess-and-check sheet identical in
behaviour to a standalone-workbook `L` sheet: the operator confirms the pre-filled
setup block, copies the `Guesses` column into the platform, and drives the loop with
the three feedback buttons (§8–§9). The three macros are unchanged and live with the
generated sheet.

Open design choices still to settle with the author:

- **Clipboard vs paste block:** the macros currently `.Copy` the guess column; the
  template rebuild could instead maintain a contiguous paste-ready block if that
  suits the live site better.
- **Hint→game mapping** when hints are sparse or out of order (see §11.4).

### 11.9 Special-case summary

| # | Situation | Behaviour |
|---|---|---|
| 1 | No hints (or none usable) | Guess Center = mean of example(s); run with no hint bounds (§11.6). |
| 2 | Hint not in `between x and y` form (e.g. "answer is odd") | Ignore that hint; treat its game as un-hinted (§11.4). |
| 3 | Example answer non-numeric (text / alphanumeric / delimited series) | Abort: **"Guess and Check only works on numeric answers"**; no sheet created (§11.3). |
| 4 | Multiple examples (`ExampleNa`, `ExampleNb`, …) | Use the mean of the examples as the example component; still 50/50 vs. hint mean (§11.6). |
| 5 | `_GCN` already exists | Create the next free `_GCN(2)`, `_GCN(3)`, … — never overwrite (§11.1). |
