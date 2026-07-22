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
     central starting guess in units of the answer's significance — **but never
     outside that game's feasible range** (a hint's `[min,max]`, never below `0`
     unless negatives are allowed, and — for un-hinted games — a range re-derived
     from the answers already solved).
  2. **Attribution search (which game).** When a batch scores `0 < k < (games
     submitted)`, you know *how many* guesses in the column were right but not
     *which*. The tool bisects the submitted set, re-testing one half at a time.
     Because the operator enters the **exact** count, each sub-test is decisive: a
     half that scores `0` is eliminated wholesale, a half that scores its own size
     is solved wholesale, and only a genuinely mixed half is split again.

The operator reports the round by **points** (what the platform shows), not a
`0/1/2+` bucket: a grid of one-click buttons for `0…7` correct, each captioned with
that many games' worth of points, plus a typed **8+** fallback. Either way the tool
receives the **exact number correct**.

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

## 6. Feedback buttons (points) + the 8+ entry

The platform reports **points**, so the buttons are captioned in **points**, not
counts. Under the header *"Number of points:"* (`E6`) sits a **4×2 grid of eight
buttons, one per number correct `0…7`**, each captioned with the points the platform
shows for that many right — `count × P` (points per game). The operator just clicks
the button whose number matches the score on screen.

| Button (count) | Caption (e.g. `P = 11`) | Placement | Bound macro |
|---|---|---|---|
| 0 | `0` | `E8:F9` | `gc_fb0` |
| 1 | `11` | `G8:H9` | `gc_fb1` |
| 2 | `22` | `I8:J9` | `gc_fb2` |
| 3 | `33` | `K8:L9` | `gc_fb3` |
| 4 | `44` | `E10:F11` | `gc_fb4` |
| 5 | `55` | `G10:H11` | `gc_fb5` |
| 6 | `66` | `I10:J11` | `gc_fb6` |
| 7 | `77` | `K10:L11` | `gc_fb7` |

Each `gc_fbK` is a one-line wrapper that calls the shared handler `gc_feedback(K)`
with its fixed count. Captions are computed once at build from `P` (`B3`); if `P`
changes afterwards, rebuild the sheet to refresh them.

**8 or more correct** happens on early scanning rounds when many games share an
answer. For that, the operator types the **points** shown into entry cell **`N8`**
(labelled *"type points:"* at `N7`, under *"8 or more?"* at `N6`) and clicks the
**`8+ pts`** button (`O8:P9`, macro `gc_fbN`). `gc_fbN` divides the points by `P` to
recover the count and hands it to `gc_feedback`. (If `P` is unknown it treats the
entry as a raw count.)

All buttons funnel into one private handler, `gc_feedback(k)`, with the exact
count `k`.

---

## 7. Working table (header at "Game Numbers" row; data below)

Columns are found by role, not by fixed row — the code locates the header row by
searching column A for the text **"Game Numbers"**, then works from the next row
down to the last populated game.

| Col | Header | Written by | Meaning |
|---|---|---|---|
| **A** | Game Numbers | formula/literal | One row per game, numbered `B4…C4`. |
| **B** | Correct Answers | macros | The confirmed answer for a solved game. **Blank = unsolved.** A numeric value here (including `0`) means solved. |
| **C** | Guesses | macros | The candidate value to enter on the platform **this** round, per game. **The whole column (blanks and all) is what you submit** — a blank game is left unanswered, i.e. scored wrong. |
| **D** | Eliminated Min | macros | Low edge of the contiguous block of values already proven wrong for this game. |
| **E** | Eliminated Max | macros | High edge of that eliminated block. The answer lies **outside** `[D,E]`; new guesses probe just past the edges (`D - sig` or `E + sig`). |
| **F** | Range Low | build + macros | Feasible **lower** bound for this game's answer. Set from a hint at build (fixed), or re-derived from solved answers for un-hinted games (adaptive). Blank = open below (the code then floors at `0`, or leaves it open when negatives are allowed). |
| **G** | Range High | build + macros | Feasible **upper** bound. Same sourcing as `F`. Blank = open above. A guess is never generated outside `[F,G]`. |
| **H** | Initial Guesses | formula | The seed guess for each game before any elimination. |
| **I** | Attribution | macros | Scratch column for the attribution search: while the tool is isolating which of several guesses are right, each candidate's value is parked here so it can be re-tested in subsets. **Non-empty anywhere in this column ⇒ the tool is mid-attribution** (rather than plain scanning). Empty during normal scanning. (The 8+ points-entry cell is `N8`, well above the table, so it never collides with this column's data.) |

**Initial-guess formulas (`H`)** — the first three data rows key off the three hint
ranges; the rest use the guess center:

```
H(row1) = IF(B7="", B10, CEILING.MATH(AVERAGE(B7:C7), B6))
H(row2) = IF(B8="", B10, CEILING.MATH(AVERAGE(B8:C8), B6))
H(row3) = IF(B9="", B10, CEILING.MATH(AVERAGE(B9:C9), B6))
H(rowN) = IF(A(rowN)="", "", B10)          ' N >= 4
```

**Range Low/High (`F`,`G`) at build** — a hinted game (one of the first three whose
`Hint N Range` in `B7:C9` parsed) gets its hint's `[min,max]` written to `F`/`G` and
never changed. Un-hinted games start blank in `F`/`G` and are (re)filled adaptively.

`B`, `C`, `D`, `E`, `I` and the adaptive part of `F`/`G` carry **no formulas** — they
are pure state written and cleared by the feedback macros.

---

## 8. Feedback handler — exact rules

All the points buttons (§6) call one private handler `gc_feedback(k)` with the
**exact count `k`** of currently-submitted guesses that scored right.

**Shared preamble.**

- `ws = ActiveSheet`; read `sig = B6` and `neg = (B11 ≠ 0)` (Negative Allowed).
- Locate header row by finding `"Game Numbers"` in column A; `firstRow = header+1`,
  `lastRow =` last non-empty cell in column A.
- Column map: A=Game, B=Correct, C=Guess, D=ElimMin, E=ElimMax, F=RangeLo,
  G=RangeHi, H=InitGuess, I=Attribution.
- **`0` is always a valid answer** — "solved" means *numeric and non-empty*, never
  "non-zero".
- After updating state, the handler **copies** the `Guesses` column to the clipboard
  so the operator can paste-submit straight away.

**Mode.** The handler first decides which search it is in:

- **Attribution mode** iff **any** cell in the `Attribution` column (`I`) is
  non-empty. Otherwise **Scanning mode**.

The "window" in either mode is simply *the set of games with a non-empty `Guess`* —
that is exactly what was submitted and scored.

### 8.1 Scanning mode

Let `m =` number of games submitted (non-empty `Guesses`). Reject `k > m` with a
message. Then:

- **`k = 0` — all wrong.** For each submitted game, fold its guess into the
  eliminated block (§8.3) and clear the guess. Then **regenerate** the scan (§8.4).
- **`k = m` — all right.** For each submitted game, copy its guess into `Correct
  Answers` and clear the guess. Then **regenerate** (§8.4) — usually the level is now
  solved and the column comes back empty.
- **`0 < k < m` — ambiguous → start attribution.** Copy every submitted guess into
  its `Attribution` cell (parking the candidates), then **hold half** (§8.3): keep
  the first half (rounded up) of the guesses active, blank the `Guesses` of the rest.
  Only the active half is submitted next round.

### 8.2 Attribution mode

The window is the active subset (non-empty `Guesses`); its parked value for each row
also sits in `Attribution`. Let `w =` window size, reject `k > w`. Then:

- **`k = 0` — the whole window is wrong.** For each window game: fold its value into
  the eliminated block (§8.3), clear its `Guess` **and** `Attribution` (it rejoins
  the scan). *Resolved.*
- **`k = w` — the whole window is right.** For each window game: copy its value into
  `Correct Answers`, clear `Guess` and `Attribution`. *Resolved.*
- **`0 < k < w` — mixed → split.** **Hold half** (§8.3): keep the first half of the
  window active, blank the `Guesses` of the rest (their `Attribution` stays, so they
  remain held candidates). *Not resolved* — the smaller window is submitted next.

After a **resolved** step, look at what candidates remain (`Attribution` non-empty):

- **Some remain →** *re-gather*: copy every remaining candidate's `Attribution` back
  into its `Guesses`, submitting them all as one window. The next count re-derives
  how many are still right, and bisection continues.
- **None remain →** attribution is finished; **regenerate** the scan (§8.4).

> **Why exact counts help.** A window that scores `0` is eliminated in one shot, and
> one that scores its full size is solved in one shot — no further probing. Only a
> genuinely mixed window is split. With the old `2+` bucket you could not tell "all
> of this half is right" from "some of it is", so those wholesale prunes were
> impossible. Re-gathering all remaining candidates after each resolution keeps the
> state to a single scratch column (no explicit stack) at the cost of a few extra
> re-test rounds; it is deliberately robust rather than round-optimal.

### 8.3 Eliminating a value / holding half

**Eliminate `v` into game `r`'s block** (`D`,`E`):

- No block yet (`D` empty) → `D = E = v`.
- Else extend the adjacent edge: if `v = D - sig` then `D = v`; if `v = E + sig` then
  `E = v`. (A non-adjacent value only ever arrives defensively; it widens the block
  to include `v`.)

**Hold half** — of the current window (non-empty `Guesses`, in row order): keep the
first `⌈n/2⌉` active; blank the `Guesses` of the rest. Nothing else moves.

### 8.4 "Regenerate the scan" subroutine

Rebuild the whole `Guesses` column for a fresh scanning round. First, refresh the
**adaptive range/centre** for un-hinted games (§8.5), and clear every `Guesses` and
`Attribution` cell. Then, for each game from `firstRow` to `lastRow`:

- If `Correct Answers` is numeric → solved, leave the guess blank.
- Else compute the game's **feasible bounds** `[lo, hi]` — `lo = F` if set, else `0`
  (or open when `neg`); `hi = G` if set, else open — and its **centre** (see §8.5),
  then:
  - **No eliminated block yet** → guess `=` the centre, clamped into `[lo, hi]`.
  - **Otherwise** → step just outside the eliminated block on whichever edge is
    **closer to the centre** (ties go low): the candidate values are `D - sig` and
    `E + sig`; keep only those still inside `[lo, hi]`; pick the closer-to-centre of
    the survivors. If **both** edges fall outside `[lo, hi]`, the game's range is
    **exhausted** and its guess is left blank.
- If, after this, every unsolved game came back blank (all exhausted), show a message
  telling the operator to widen `Range Low`/`Range High` or check `Negative Allowed`.

This replaces the old random-edge probe: the edge choice is **deterministic**
(closest-to-centre, ties low) and always **clamped** to the feasible range, so the
search never wanders below `0` (unless negatives are allowed) or outside a hint.

### 8.5 Adaptive range and centre (un-hinted games)

Most games have no hint of their own, so their search is bounded only by what has
already been solved. On every scan regeneration, once **≥ 2** games are solved, with
`minS`/`maxS`/`spread = maxS − minS` over the solved answers:

- **Range** (written to `F`/`G` of every un-hinted, unsolved game):
  `half = max(3·spread, 20·sig)`, `Range Low = minS − half` (floored at `0` unless
  negatives are allowed), `Range High = maxS + half`. The margin is deliberately
  generous — it exists to stop runaway scanning, not to gamble that the answer sits
  in the solved cluster.
- **Centre** (used only to seed the first guess and to pick the closer edge): the
  **mean of the solved answers** for un-hinted games (once ≥ 2 are solved), else the
  game's `Initial Guesses (H)`. Hinted games always centre on their own `H` (the hint
  midpoint) and keep their fixed hint range.

Hinted games are never touched by the adaptive step; their `F`/`G` stay pinned to the
hint. All of `F`/`G` remain operator-editable — clearing them re-opens a game's range.

---

## 9. Operator workflow (the loop)

1. Open the `L` sheet for the level; confirm `B2` is the level number and the setup
   block populated (games, points, center, significance) from `case data`. Fill in
   any known `Hint 1/2/3` ranges (`B7:C9`) if you have them.
2. The `Guesses` column starts at the initial guesses. **Copy it and paste into the
   platform's answer boxes for that level; submit.**
3. Read the level's **points** off the scoreboard.
4. Report it: click the **points button** whose caption matches (buttons for `0…7`
   correct). If you scored more points than any button shows (8+ correct), type the
   points into cell `N8` and click **`8+ pts`**. The handler updates state and puts
   the **next** guess column on the clipboard.
5. Paste, submit, repeat. Games move into `Correct Answers` as they're pinned down;
   when every game is solved the `Guesses` column is empty and `Correct Answers` is
   full.

> When a round scores `2`+, the tool enters an **attribution** pass: it re-submits
> smaller and smaller subsets (the `Guesses` column will have fewer entries) to work
> out which games were the right ones. Keep pasting-and-reporting exactly as before —
> the `Guesses` column always holds precisely what to submit next.

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
  Extension only happens when a value is exactly one `sig` past an edge — the
  contiguous scan guarantees that, so eliminations never "jump".
- **`[RangeLo, RangeHi]` (`F`,`G`) is the feasible range** — the outer wall the scan
  never crosses. A hinted game's is fixed; an un-hinted game's is adaptive. Blank =
  open on that side (still floored at `0` unless negatives are allowed).
- **`Attribution` (`I`) is the mode flag *and* the candidate scratch.** Non-empty
  anywhere ⇒ mid-attribution. A single scratch column (plus re-gathering all
  remaining candidates after each resolution) replaces the old two-bucket scheme and
  needs no stack.
- **The count is exact.** `k = 0`, `k = m` (all), and `0 < k < m` are genuinely
  different branches — the "all right" and "all wrong" wholesale steps depend on
  knowing the exact number, not a `2+` bucket.
- **"Hold half" rounds up** (`⌈n/2⌉` stays active). One rule, used everywhere a
  window is split.
- **Header is found by text**, and the row range by the last populated game number,
  so the table can start on any row and be any length.
- **Negative Allowed (`B11`)** bounds the downward search: when `0`, no guess is ever
  generated below `0`; when `1`, the low side is open (down to the adaptive/hint
  bound). This clamp is now enforced in the scan.
- **Formatting is cosmetic only.** Fills, borders, and widths carry no state; the
  logic reads values, never colours. A generated sheet re-styled by hand still works.

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

`_GCN` (or `_GCN(k)`) is a standalone guess-and-check sheet: the operator confirms the
pre-filled setup block, copies the `Guesses` column into the platform, and drives the
loop with the points buttons + the 8+ points-entry cell (§6, §8–§9). The
build also runs one cosmetic formatting pass (§11.10).

Open design choices still to settle with the author:

- **Adaptive-range margin.** The un-hinted feasible range is `solved-cluster ±
  max(3·spread, 20·sig)` (§8.5) — generous on purpose. If a level's answers turn out
  to spread far wider than the first few solved, a game can exhaust its range and the
  tool prompts to widen `F`/`G`. The multiplier is a tunable heuristic, not a proven
  bound.
- **Hint→game mapping** when hints are sparse or out of order (see §11.4).

### 11.10 Sheet formatting

After the layout and first guesses are written, a cosmetic pass styles the sheet:
a title row, bold shaded setup labels, a bordered setup value block, a shaded/bold
table header, light borders over the data grid, a green `Correct Answers` column, a
**highlighted `Guesses` column** (the one to copy) with a darker header cap, a
bordered yellow points-entry cell (`N8`) for the 8+ fallback, and uniform column
widths (`A:L`, so the `0…7` button grid reads evenly). None of it
carries state — it is purely to make the working sheet readable at a glance.

### 11.9 Special-case summary

| # | Situation | Behaviour |
|---|---|---|
| 1 | No hints (or none usable) | Guess Center = mean of example(s); run with no hint bounds (§11.6). |
| 2 | Hint not in `between x and y` form (e.g. "answer is odd") | Ignore that hint; treat its game as un-hinted (§11.4). |
| 3 | Example answer non-numeric (text / alphanumeric / delimited series) | Abort: **"Guess and Check only works on numeric answers"**; no sheet created (§11.3). |
| 4 | Multiple examples (`ExampleNa`, `ExampleNb`, …) | Use the mean of the examples as the example component; still 50/50 vs. hint mean (§11.6). |
| 5 | `_GCN` already exists | Create the next free `_GCN(2)`, `_GCN(3)`, … — never overwrite (§11.1). |
