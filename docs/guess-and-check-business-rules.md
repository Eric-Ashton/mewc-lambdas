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
  1. **Value search (per game).** Each unsolved game tries its next candidate. Order:
     values already **confirmed elsewhere on the level** (most frequent first) — since
     a level's answers often repeat, testing a known value across all games is the
     fastest route — then an outward scan from a centre that re-derives from solved
     answers (mode/median). Hinted games are walled to their
     hint; un-hinted games are floored at `0` (unless negatives are allowed) and
     otherwise **unbounded**, so the search never declares a value impossible.
  2. **Attribution search (which game).** When a batch scores `0 < k < (games
     submitted)`, the tool bisects the submitted set — but it **carries the exact
     count through the split**: a parent group of known count `k` splits into a tested
     half (count `a`) and a sibling whose count is `k − a` *for free*. A sheet-persisted
     stack of `(group, known-count)` means all-right and all-wrong sub-sets are settled
     **inside Excel** with no re-test; only genuinely mixed sub-sets cost a round.

The operator reports each round by the **absolute points** the platform shows. Because
confirmed answers stay in the submitted column, points bank on the leaderboard as you
go, and the `0…7` buttons are captioned with the exact platform total each round (plus
a typed `8+` fallback and an `Undo`).

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

**Significance rule (`B6`) — the FINEST step the samples justify, never coarser than
`1`.** Take, per sample, the smallest power of ten `≤ 1` that divides it exactly
(`1` for any integer, `0.1`/`0.01`/… for decimals), and use the **finest** (smallest)
across all samples:

- all-integer samples → **`sig = 1`** — e.g. `20`, `439`, `90` all give `1`.
- any decimal sample → the smallest decimal unit — `0.05 → 0.01`.

> **Important (was a bug):** never infer a step *coarser* than `1` just because a
> sample happens to divide by 10 (e.g. `20`). An earlier version returned the
> *largest* dividing power of ten, so a level whose example and hints were all round
> numbers (say example `20`, hint "between `90` and `100`") got `sig = 10` and the
> search stepped `90, 100, 110…`, never testing `91–99` — an unsolvable level. Getting
> `sig` too *large* silently loses answers; too *small* only costs rounds, so we bias
> to `1`. Only override to a coarser step when the case explicitly says answers are
> multiples of 10/100/…

**Row 12** is a diagnostics row: `A12/B12/C12/K12` are `=COUNT(…14:…1000)` — the live
count of games, solved answers, active guesses, and parked attribution candidates.

---

## 6. Feedback buttons — captioned with the platform's ABSOLUTE score

Confirmed answers stay in the submitted column (§7, `Submit`), so the platform score
always includes them: **`points = (confirmed + active-hits) × P`**, where *active
hits* = how many of the guesses still being tested this round are right. The buttons
report **active hits** but are **captioned with the absolute score** the platform will
show, so the operator just clicks the number that matches the screen.

Under *"Points on the platform…"* (`E6`) is a **4×2 grid of eight buttons for `0…7`
more correct**, plus an **`8+ pts`** fallback and an **`Undo`** button:

| Button | Reports active hits | Caption (dynamic) | Placement | Macro |
|---|---|---|---|---|
| 0…7 | `0…7` | `(confirmed + j) × P` | `E8…L11` (4×2) | `gc_fb0…gc_fb7` |
| `8+ pts` | typed | — | `O8:P9` | `gc_fbN` |
| `Undo` | — | — | `O10:P11` | `gc_undo` |
| `Re-evaluate` | typed | — | `Q8:R11` | `gc_reeval` (§8.7) |

- **Captions are recomputed after every round** (via `gc_recaption`), because
  `confirmed` grows as games are solved — a button for "`j` more" always shows the
  exact platform total `(confirmed + j) × P`. (Named form-control buttons `gc_btn0…7`
  hold the captions; they can't hold formulas, hence the refresh.)
- **`8+ pts`** — for 8 or more more-correct (common on the first scan): type the
  **absolute points** into cell **`N8`** and click. `gc_fbN` computes
  `active-hits = points/P − confirmed`, rejecting a points value that isn't a whole
  multiple of `P`.
- **`Undo`** reverts the last feedback (§8.6).

All buttons live in the `gc_buttons` module (`Option Private Module`, so they stay out
of Alt+F8) and forward the **active-hit count** to the one public handler
`gc_feedback(activeHits)` in `guess_and_check` (its required argument also keeps it out
of Alt+F8). `create_gc_sheet` is the only Alt+F8-visible entry point.

---

## 7. Working table (header at "Game Numbers" row; data below)

Columns are found by role, not by fixed row — the code locates the header row by
searching column A for the text **"Game Numbers"**, then works from the next row
down to the last populated game.

| Col | Header | Written by | Meaning |
|---|---|---|---|
| **A** | Game Numbers | literal | One row per game, numbered `B4…C4`. |
| **B** | Correct Answers | macros | The confirmed answer for a solved game. **Blank = unsolved.** Numeric (incl. `0`) = solved. |
| **C** | Guess | macros | The candidate value being **tested** this round, per unsolved game. Blank for solved games and for held/idle games. |
| **D** | Submit  *(copy this)* | **formula** | `=IF(ISNUMBER(B),B,IF(C="","",C))` — confirmed answer if solved, else the active guess, else **blank** (a bare `C` reference would coerce a blank to `0`, and `0` is a valid answer we must not submit for a held game). **This is the column the operator copies**, so confirmed answers are always resubmitted and the leaderboard banks points continuously. |
| **E** | Elim Min | macros | Low edge of the contiguous block of values already proven wrong for this game. |
| **F** | Elim Max | macros | High edge of that block. |
| **G** | Tried Extras | macros | Comma-list of **non-contiguous** values already tried (e.g. a priority-value probe of `12` while `9–11` are still untested). Needed because `[Elim Min, Elim Max]` alone can't represent gaps. |
| **H** | Hint Lo | build | **Hard** lower bound — a hint's min. Blank = no hard lower wall. Never changes. |
| **I** | Hint Hi | build | **Hard** upper bound — a hint's max. Blank = no hard upper wall. Never changes. |
| **J** | Initial Guess | formula | The seed centre for each game before any solved data. |
| **K** | Attribution | macros | Parked candidate value during attribution. **Non-empty anywhere ⇒ mid-attribution** (else scanning). |
| **L** | Grp | macros | Attribution group tag: a positive id = member of that pending group on the stack; `-1` = the half submitted now; `-2` = its held sibling; blank = not a candidate. |

A value is **"tried"** for a game iff it is inside `[Elim Min, Elim Max]` **or** listed
in `Tried Extras`. The answer is whichever untried value the search reaches.

**Initial-guess formulas (`J`)** — first three rows use their hint midpoint, the rest
the guess centre (`B10`); this seeds round 1 only (later rounds re-centre on solved
data, §8.5):

```
J(row1) = IF(B7="", B10, CEILING.MATH(AVERAGE(B7:C7), B6))   ' rows 2,3 use B8/B9
J(rowN) = IF(A(rowN)="", "", B10)                            ' N >= 4
```

**Hard bounds (`H`,`I`) at build** — a hinted game (first three, whose `Hint N Range`
parsed) gets its hint's `[min,max]`; these are **walls the search never crosses**.
Un-hinted games leave `H`/`I` blank: they are **unbounded above** and floored at `0`
(unless *Negative Allowed*), and the "learned" cluster only affects search *order*, so
they **never declare a value permanently impossible** and never "run out of range".

**Resolver state** lives off-screen (hidden columns `T:BB`): scalars in `T1:T4` (stack
depth, next group id, parent count, round) and the LIFO group stack in `U:V`; `AN:BB`
hold the one-level undo snapshot.

`B`, `C`, `E`, `F`, `G`, `K`, `L` carry **no formulas** — pure state written by the
handler. `D` and `J` are formulas; `A`, `H`, `I` are set once at build.

---

## 8. Feedback handler — exact rules

The buttons (§6) call one public handler `gc_feedback(activeHits)` with the **exact
number of the currently-tested guesses that were right** (confirmed answers excluded).

**Shared preamble.**

- `ws = ActiveSheet`; read `sig = B6`, `neg = (B11 ≠ 0)`; locate the header by finding
  `"Game Numbers"` in column A; `firstRow = header+1`, `lastRow =` last game in A.
- Column map: A=Game, B=Correct, C=Guess, D=Submit, E=ElimMin, F=ElimMax, G=Tried,
  H=HintLo, I=HintHi, J=Init, K=Attribution, L=Grp.
- **`0` is always a valid answer** — "solved" means *numeric and non-empty*.
- **Validate first (no state change on a bad number).** Let `sub =` count of active
  guesses (`C` non-empty). Reject unless `0 ≤ activeHits ≤ sub`; in attribution also
  reject if `activeHits > parentK` or `parentK − activeHits > (sibling size)`. A
  mis-read score is refused rather than corrupting state.
- **Snapshot then act.** Copy the mutable state to the hidden backup (one-level undo,
  §8.6) *before* changing anything.
- After acting, refresh `D` (a formula, automatic), **re-caption** the buttons, and
  **copy the `Submit` column** to the clipboard.

**Mode.** **Attribution** iff any `Attribution` (`K`) cell is non-empty, else
**Scanning**. The submitted set is *the games with a non-empty `Guess`*.

### 8.1 Scanning step

Let `m =` games submitted. With `k = activeHits`:

- **`k = 0`** — fold every guess into its tried state (§8.3), clear guesses,
  **regenerate** the scan (§8.4).
- **`k = m`** — copy every guess into `Correct Answers`, clear, **regenerate**.
- **`0 < k < m`** — **park the whole submitted set as one attribution group** of known
  count `k`: write each guess into its `Attribution`, tag all rows with a fresh group
  id, push `(id, k)` on the stack, clear the guesses. Then **pump** (§8.2).

### 8.2 Attribution step — carrying the known count (no re-gather)

Attribution resolves a **stack of groups, each with a KNOWN correct-count**, so a
sub-set that is all-right or all-wrong is settled *inside Excel* with no extra
submission. Only a genuinely mixed group costs a website round.

**On feedback** (the submitted half — the rows tagged `Grp = -1` — scored `a`):

1. The held sibling (`Grp = -2`) then contains `parentK − a` correct **for free**
   (`parentK` is the split's parent count, kept in state).
2. **Absorb** each of the two halves by its now-known count (§ below).
3. **Pump** the stack.

**Absorb a set of rows whose known count is `c`:**

- `c = 0` → all wrong: eliminate each (§8.3), free them.
- `c = size` → all right: copy each into `Correct Answers`.
- `0 < c < size` → still mixed: give it a fresh group id and **push `(id, c)`**.

**Pump** — repeat until a group must be tested or the stack is empty:

- **Stack empty** → attribution done; **regenerate** the scan (§8.4).
- **Pop** the top `(id, c)`. Its members are the rows tagged `id`. (`c = 0` / `c =
  size` are settled immediately as above — but only mixed groups are ever pushed.)
- **Mixed** → split the members: first `⌈n/2⌉` become the active half (`Grp = -1`,
  `Guess = Attribution`, i.e. submitted); the rest the sibling (`Grp = -2`, guess
  blank). Record `parentK = c`. **Stop** — the operator submits and reports `a`.

So a scan of `M` games scoring `k` parks once and thereafter each **mixed** group
takes exactly one submission to split; all-right/all-wrong halves are free. This is
the "carry the count" method — roughly half the website rounds of the old re-gather.

> **Worked step.** Scan of 20 scores `k = 2`. Park all 20 as `(g1, 2)`. Pump splits
> `g1` → submit rows 1–10. They score `a = 1` ⇒ rows 11–20 hold `2 − 1 = 1` (free).
> Absorb: rows 1–10 → `(g2, 1)` pushed; rows 11–20 → `(g3, 1)` pushed. Pump pops `g3`,
> splits it, submits rows 11–15… and so on, never re-testing a settled subset.

### 8.3 Eliminating a value (contiguous block + tried-extras)

**Eliminate `v` for game `r`:**

- No block yet → `Elim Min = Elim Max = v`.
- `v` adjacent to an edge (`= ElimMin − sig` or `= ElimMax + sig`) → extend that edge,
  **then swallow any `Tried Extras` value now adjacent** (so a run tried out of order —
  e.g. a priority probe of `20` before the block scanned up to it — collapses back into
  one block rather than leaving `9–19` + extras `20–26`).
- `v` already inside `[ElimMin, ElimMax]` → nothing.
- otherwise (a non-contiguous probe, e.g. a priority value away from the block) →
  append `v` to **`Tried Extras`**.

A value counts as **tried** if it's in the block **or** the extras list — both the
outward scan and the priority check consult both. So `Tried Extras` values are just as
eliminated as the block; they're only listed separately when they aren't (yet)
contiguous with it.

### 8.4 "Regenerate the scan" subroutine

Clear every `Guess`, `Attribution`, and `Grp`; compute the **centre** and **priority
values** from the solved answers (§8.5). Then for each unsolved game, set its `Guess`
to the first candidate that is **untried and in-bounds**, tried in this order:

1. **Priority values** — values already confirmed elsewhere on the level (most
   frequent first, ties low). *This is the big lever on clustered levels: once `12` is
   a known answer, every unsolved game tests `12` next, and §8.2 attributes the whole
   cluster cheaply.*
2. **Outward scan** from the (grid-snapped) centre — nearest value to the centre,
   ties low — skipping tried values.

**Bounds.** Hinted games are walled to `[Hint Lo, Hint Hi]`. Un-hinted games are
floored at `0` (unless `neg`) and **unbounded above** — the outward scan simply keeps
going, so an un-hinted game never "exhausts". Only a **hinted** game can run out (its
whole hint range tried) — then its guess is blank and a message flags it (bad hint or
wrong `sig`). The learned cluster changes *order*, never possibility.

### 8.5 Centre and priority values (from solved answers)

Recomputed every regeneration from the confirmed answers (`B`):

- **Centre** = the **mode** of solved answers if any value repeats, else their
  **median**; with nothing solved yet, the build centre `B10` (the example/hint blend
  — a useful level-wide prior for round 1). Solved data supersedes the seed as soon as
  it exists. Hinted games still centre on their own hint midpoint.
- **Priority values** = the **distinct** solved values, most-frequent first (ties
  low). (The example answer is *not* a priority value — it feeds the centre only, so
  round 1 probes the centre rather than a single other game's answer.)

> This keeps ChatGPT's intent (drive centring and probe order from real solved data)
> while keeping the hint-informed seed for round 1 — in the Level-5 recording the
> blended seed (`10`) sat on the answers (`11–13`) whereas the bare example (`4`)
> would have been worse.

### 8.6 Undo and recovery

Every `gc_feedback` **snapshots the mutable state first** (values of `B,C,E:L` and the
resolver scalars/stack) into hidden backup columns — one level deep. The **`Undo`**
button (`gc_undo` → `gc_apply_undo`) restores that snapshot, so a single mis-clicked
or mis-read score is fully reversible. Combined with the up-front validation (§8
preamble), one stale leaderboard reading can't silently corrupt a long solve.

### 8.7 Re-evaluate (recover from a deeper error)

`Undo` only steps back one round. If a mis-click happened several rounds ago and left
a **wrongly-confirmed answer**, the honest platform score drops below what the sheet
expects and *no button matches*. **`Re-evaluate`** (`gc_reeval` → `gc_do_reeval`)
recovers from a single fact: type the **current platform points** into `N8` and click.
It then:

1. Reads the whole **current `Submit` column** (confirmed answers + active guesses) and
   its true count `truePoints / P`.
2. **Forgets the confirmed/guessed distinction** — un-confirms everything, abandons the
   attribution stack — but **keeps** each game's `Elim`/`Tried` history (so no known-
   wrong value is re-tried).
3. Re-parks the entire submission as **one attribution group of that known count** and
   resolves it by the normal carry-count bisection (§8.2).

Because the count is ground truth, a wrongly-confirmed answer lands in a sub-set that
scores short and is dropped (its value folded into `Tried`); the genuinely-correct
answers are re-confirmed. It "starts from the knowledge of that one score" rather than
restarting the level. (The platform total dips while the confirmed set is re-verified,
then climbs back as answers re-confirm — the price of certainty after an error.)

---

## 9. Operator workflow (the loop)

1. Confirm the setup block (level, points, significance, hints). The build already
   put the first submission on the clipboard.
2. **Copy the `Submit` column (`D`) and paste it into the platform; submit.** `Submit`
   always carries your confirmed answers plus the current guesses, so the leaderboard
   banks every solved game as you go.
3. Read the level's **points** off the scoreboard.
4. **Click the button whose caption equals that number** (buttons `0…7` more correct,
   re-captioned to the absolute total each round). If the score is higher than any
   button shows, type the **points** into `N8` and click **`8+ pts`**. The handler
   updates state and puts the next `Submit` column on the clipboard. Mis-clicked?
   **`Undo`** reverts it.
5. Paste, submit, repeat. Games flow into `Correct Answers`; when all are solved the
   `Submit` column is exactly your answer key and the leaderboard already has them.

> When a round leaves some games ambiguous, the tool enters an **attribution** pass:
> the `Submit` column carries your banked answers plus a shrinking subset of test
> guesses. Keep pasting-and-reporting exactly as before — `Submit` always holds
> precisely what to send next, and the button captions always match the platform total.

---

## 10. Conventions and edge cases (must preserve on rebuild)

- **`0` is a valid answer.** Every "is this solved / is this a real value" test must
  be *numeric-and-present*, not *non-zero*. This is the single most common way a
  reimplementation breaks.
- **Significance `sig` is the step unit.** All probing moves by `± sig`, and initial
  guesses are rounded with `CEILING.MATH(…, sig)`. Non-integer answers work if `sig`
  is set correctly (e.g. `0.01`).
- **"Tried" = block ∪ extras.** A value is ruled out for a game if it is inside
  `[Elim Min, Elim Max]` or listed in `Tried Extras`. The contiguous block compresses
  the swept-from-centre region; the extras list holds one-off priority-value probes.
  The search returns the nearest untried in-bounds value.
- **Hard bounds vs. order.** `Hint Lo`/`Hint Hi` (`H`,`I`) are the only *walls* — set
  once from a hint. Everything else (the learned centre, priority values, the
  unbounded un-hinted scan) shapes the *order* of guesses, never forbids a value. So
  an un-hinted game can't "exhaust"; only a hinted game can (its hint range fully
  tried), which flags a bad hint / wrong `sig`.
- **`Attribution` (`K`) is the mode flag; `Grp` (`L`) is the stack tag.** Non-empty
  `Attribution` ⇒ mid-attribution. Resolution is a proper LIFO stack of
  `(group, known-count)` (state in hidden `T:V`): the exact count is *carried* into
  each split, so all-right / all-wrong sub-sets never cost a re-test. `⌈n/2⌉` stays
  active on every split.
- **Confirmed answers are always submitted** (`Submit = IF(ISNUMBER(B),B,C)`), so the
  platform total is `(confirmed + active-hits) × P`; the handler and buttons work in
  that absolute frame, and the leaderboard banks continuously.
- **Every feedback is validated and snapshotted** — an impossible count is refused
  (§8 preamble) and `Undo` reverts the last applied round.
- **`0` is a valid answer**; **`sig` is the step unit** (all probing moves by `± sig`);
  **header is found by text** so the table can start on any row.
- **Formatting is cosmetic only.** Fills, borders, widths carry no state; the logic
  reads values, never colours.

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

Compute `sig` by the **§5 rule — the finest step the samples justify, never coarser
than `1`**: per sample take the smallest power of ten `≤ 1` that divides it exactly
(integers → `1`), and use the finest across the example answer(s) and usable hint
`min`/`max`. E.g. `87, 90, 100, 350 → 1`; `20, 90, 100 → 1` (**not** `10`); `0.05,
0.10 → 0.01`. Write it to `B6`, editable — override to a coarser step only when the
case explicitly says answers are multiples of 10/100/…. (See the §5 note on why a
too-coarse `sig` silently loses answers.)

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
pre-filled setup block, copies the **`Submit` column** into the platform, and drives
the loop with the points buttons + the 8+ entry + `Undo` (§6, §8–§9). The build also
runs one cosmetic formatting pass (§11.10).

Open design choices still to settle with the author:

- **Priority-value probing** front-loads confirmed answer values across all games
  (§8.4–8.5); if a level's answers are all distinct it wastes a few probes (each
  recorded in `Tried Extras`) before the outward scan takes over — a cheap bet given
  how often MEWC answers repeat.
- **Hint→game mapping** when hints are sparse or out of order (see §11.4).

### 11.10 Sheet formatting

After the layout and first guesses are written, a cosmetic pass styles the sheet:
a title row, bold shaded setup labels, a bordered setup value block, a shaded/bold
table header, light borders over the data grid, a green `Correct Answers` column, a
**highlighted `Submit` column** (the one to copy) with a darker header cap, a bordered
yellow points-entry cell (`N8`), uniform column widths (`A:L`), and the resolver /
backup columns (`T:BB`) hidden. None of it carries state — it is purely to make the
working sheet readable at a glance.

### 11.9 Special-case summary

| # | Situation | Behaviour |
|---|---|---|
| 1 | No hints (or none usable) | Guess Center = mean of example(s); run with no hint bounds (§11.6). |
| 2 | Hint not in `between x and y` form (e.g. "answer is odd") | Ignore that hint; treat its game as un-hinted (§11.4). |
| 3 | Example answer non-numeric (text / alphanumeric / delimited series) | Abort: **"Guess and Check only works on numeric answers"**; no sheet created (§11.3). |
| 4 | Multiple examples (`ExampleNa`, `ExampleNb`, …) | Use the mean of the examples as the example component; still 50/50 vs. hint mean (§11.6). |
| 5 | `_GCN` already exists | Create the next free `_GCN(2)`, `_GCN(3)`, … — never overwrite (§11.1). |
