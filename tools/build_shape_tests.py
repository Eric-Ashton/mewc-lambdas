#!/usr/bin/env python3
"""Build (or rebuild) the `shape_tests` sheet in the test workbook.

What it tests
-------------
Every lambda test case on the per-lambda sheets is re-run with its arguments
transformed in ways that should NOT change the answer, and the result is compared
with the original. One row per (case, family):

  tocol    every single-value argument wrapped in TOCOL(...)  -> a 1x1 ARRAY
           (what TOCOL(x), {x} or a one-element FILTER/INDEX/TEXTSPLIT result hands
           a lambda; SEQUENCE/MAKEARRAY/REGEXEXTRACT/IF lift over it unless the
           lambda collapses it - z_scalar / z_opt / z_num / CONCAT)
  cellref  every literal argument moved into a helper cell and referenced
  textnum  every numeric argument supplied as TEXT ("5" / D6&"")
           (numbers stored as text are everywhere in case data; z_num / z_opt
           convert them for arguments that are numbers by nature)
  array    every multi-cell range argument replaced by a computed ARRAY of the
           same values (CHOOSEROWS) - a FILTER/TAKE result instead of a range
  blank    every omitted optional argument pointed at an EMPTY cell
           (z_opt: blank or "" means "use the default")

The Match cell recomputes BOTH the original formula and the variant (cell
references qualified to the source sheet) and PASSes when they agree in shape and
every element (errors compare by ERROR.TYPE).

Deliberately NOT transformed, because the change would legitimately alter meaning:
  * tocol/cellref/array on range-reference arguments whose POSITION matters -
    ad_to_cont's arr, arrow_trip's rng, cont_to_ad's arr, route_cost's cost_map,
    arrow_move's board_rng, neighbors_list's boundary_rng, a cell or range
    reference passed as move_rng's ad_or_rng;
  * textnum on values that are LOOKED UP in data or carried as data ("2" and 2 must
    stay distinct) - exists, xrank, xcountif, cont_to_ad, exlookup's lookup/if_not_found,
    neighbors_count's val, set_d/set_i/set_u operands, label_lookup labels,
    update_arr's value, scan2/scan3 start values, shift_arr's fill, xfilter's arr,
    and a case whose point is that an invalid input passes through unchanged;
  * array on a range that contains an EMPTY cell (an array turns it into 0);
  * blank on exlookup's if_not_found (a blank there IS the fallback value).

The hand-written `round_trips` sheet (chains of lambdas that must be identities) is
harvested too: its families transform the arguments of the OUTERMOST library call.

Usage
-----
    py tools/build_shape_tests.py "<src.xlsm>" "<out.xlsm>"

Re-run it whenever test cases are added or changed; it rewrites the sheet's rows
(creating the sheet, its roll-up row and tab position on first use). Then open the
workbook in Excel, run sync_all / lambda_update as usual, and check the tally.
"""
from __future__ import annotations

import glob
import json
import os
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from xlsm_edit import apply_edits, validate  # noqa: E402

SHEET = "shape_tests"
TEMPLATE_SHEET = "count_substring"          # layout donor for a fresh sheet
NON_LAMBDA_SHEETS = {"Prep", "Lamb", "lambda_tests", "vba_tests", SHEET}
EXTRA_SHEETS = {"round_trips"}               # hand-written chain tests: transform the OUTERMOST library call
LAMBDA_DIR = HERE.parent / "lambdas"

# (lambda, 1-based argument index) -> arguments the families must leave alone
POSITION_ARGS = {("ad_to_cont", 2), ("arrow_trip", 1), ("cont_to_ad", 2), ("route_cost", 1),
                 ("arrow_move", 4), ("neighbors_list", 3)}
DATA_ARGS = {("exists", 2), ("xrank", 1), ("xcountif", 2), ("cont_to_ad", 1), ("exlookup", 1), ("exlookup", 4),
             ("neighbors_count", 2), ("set_d", 1), ("set_d", 2), ("set_i", 1), ("set_i", 2), ("set_u", 1), ("set_u", 2),
             ("label_lookup", 2), ("label_lookup", 3), ("update_arr", 4), ("freq_table", 1), ("make_change", 1),
             ("biggest", 1), ("biggest", 2), ("smallest", 1), ("smallest", 2), ("textbetween", 2), ("textbetween", 3),
             ("count_substring", 1), ("count_substring", 2), ("scan2", 1), ("scan2", 2), ("scan3", 1), ("scan3", 2), ("scan3", 3),
             ("shift_arr", 3), ("xfilter", 1)}
BLANK_SKIP = {("exlookup", 4)}
# cases where the lambda PASSES AN INVALID INPUT THROUGH unchanged, so text in -> text out is correct
TEXTNUM_SKIP_CASES = {("str_ad", "H10")}
REF_ONLY_FIRST_ARG = {"move_rng"}            # a bare cell/range reference as arg 1 is a position

WRAPPABLE = re.compile(r'^\s*(\$?[A-Z]{1,3}\$?\d+|"(?:[^"]|"")*"|-?\d+(?:\.\d+)?)\s*$')
CELL_REF = re.compile(r'^\s*\$?[A-Z]{1,3}\$?\d+\s*$')
RANGE_ARG = re.compile(r'^\s*\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+\s*$')
NUM_LIT = re.compile(r'^\s*-?\d+(?:\.\d+)?\s*$')
STR_LIT = re.compile(r'^\s*"((?:[^"]|"")*)"\s*$')
REF_TOKEN = re.compile(r"(?<![A-Za-z_!'.$])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Za-z_(\d])")
HELPER_COLS = ["H", "I", "J", "K", "L", "M", "N", "O"]   # cellref helper values, same row
BLANK_REF = f"{SHEET}!$Z$1"                             # never written -> empty


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def unescape(s: str) -> str:
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


def split_args(s: str) -> list[str]:
    out, depth, cur, inq = [], 0, "", False
    for ch in s:
        if ch == '"':
            inq = not inq
            cur += ch
            continue
        if inq:
            cur += ch
            continue
        if ch in "({[":
            depth += 1
        if ch in ")}]":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def find_call(formula: str, name: str):
    """(start, end) of the first `name(...)` call whose name is a whole word."""
    i = -1
    while True:
        i = formula.find(name + "(", i + 1)
        if i < 0:
            return None
        if i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "_."):
            break
    j = i + len(name) + 1
    depth, inq = 1, False
    while j < len(formula) and depth > 0:
        ch = formula[j]
        if ch == '"':
            inq = not inq
        elif not inq:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
        j += 1
    return i, j


def qualify(formula: str, sheet: str) -> str:
    """Prefix every bare cell reference outside quotes with 'sheet'!"""
    out, inq, buf = [], False, ""
    for ch in formula:
        if ch == '"':
            out.append(buf if inq else REF_TOKEN.sub(lambda m: f"'{sheet}'!{m.group(1)}", buf))
            buf = ""
            inq = not inq
            out.append(ch)
            continue
        buf += ch
    out.append(buf if inq else REF_TOKEN.sub(lambda m: f"'{sheet}'!{m.group(1)}", buf))
    return "".join(out)


def sheet_part(z: zipfile.ZipFile, wbx: str, rels: str, name: str) -> str:
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wbx).group(1)
    m = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, rels).group(0)
    return "xl/" + re.search(r'Target="([^"]+)"', m).group(1).replace("/xl/", "").lstrip("/")


def signatures() -> dict:
    """lambda -> (total params, optional params) from the .lambda files"""
    sig = {}
    for f in glob.glob(str(LAMBDA_DIR / "*.lambda")):
        name = os.path.basename(f)[:-7]
        txt = open(f, encoding="utf-8").read()
        s = txt.split("=== SIGNATURE ===")[1].split("===")[0].strip()
        inner = s[s.index("(") + 1:s.rindex(")")]
        params = [p.strip() for p in split_args(inner)] if inner.strip() else []
        sig[name] = (len(params), sum(1 for p in params if p.startswith("[")))
    return sig


def harvest(src: str):
    z = zipfile.ZipFile(src)
    wbx = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    order = re.findall(r'<sheet [^>]*name="([^"]+)"', wbx)
    lam_sheets = [n for n in order if n not in NON_LAMBDA_SHEETS and not n.startswith("zz_")]   # includes EXTRA_SHEETS
    sig = signatures()
    lambda_names = set(sig)
    rows = []   # (sheet, cell, family, original_raw, variant_raw, helpers)
    for sn in lam_sheets:
        x = z.read(sheet_part(z, wbx, rels, sn)).decode()
        numeric_cells, filled_cells = set(), set()
        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)>(.*?)</c>', x, re.S):
            addr, attrs, inner = m.groups()
            if "<v>" in inner or "<is>" in inner or "<f" in inner:
                filled_cells.add(addr)
            if 't="' not in attrs and "<f" not in inner and re.search(r"<v>-?[\d.E+]+</v>", inner):
                numeric_cells.add(addr)

        def range_has_blank(a):
            """True when any cell of the A1:B2 range is empty on the source sheet (an array would turn it into 0)."""
            m_ = re.match(r'\s*\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)', a)
            c1, r1, c2, r2 = col_num(m_.group(1)), int(m_.group(2)), col_num(m_.group(3)), int(m_.group(4))
            for rr_ in range(min(r1, r2), max(r1, r2) + 1):
                for cc_ in range(min(c1, c2), max(c1, c2) + 1):
                    if f"{col_letters(cc_)}{rr_}" not in filled_cells:
                        return True
            return False

        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)>(.*?)</c>', x, re.S):
            addr, _attrs, inner = m.groups()
            fm = re.search(r"<f[^>]*>(.*?)</f>", inner, re.S)
            if not fm:
                continue
            raw = unescape(fm.group(1))
            if addr in ("B1", "B2", "I1", "S1") or raw.startswith(("COUNTIF(", "IF(", "HYPERLINK(")):
                continue
            if sn in EXTRA_SHEETS:
                # chain test: pick the library call that starts earliest in the formula
                found = [(find_call(raw, nm), nm) for nm in lambda_names if nm + "(" in raw]
                found = [(sp, nm) for sp, nm in found if sp]
                if not found:
                    continue
                span, call_name = min(found, key=lambda p: p[0][0])
            else:
                span, call_name = find_call(raw, sn), sn
            if not span:
                continue
            i, j = span
            args = split_args(raw[i + len(call_name) + 1:j - 1])
            fn = call_name   # the lambda whose arguments are transformed (== sn except on EXTRA_SHEETS)

            def rebuild(new_args):
                return raw[:i] + fn + "(" + ",".join(new_args) + ")" + raw[j:]

            def is_position(k, a):
                return (fn, k + 1) in POSITION_ARGS or (
                    fn in REF_ONLY_FIRST_ARG and k == 0 and (CELL_REF.match(a) or RANGE_ARG.match(a)))

            # tocol
            new, touched = list(args), False
            for k, a in enumerate(args):
                if WRAPPABLE.match(a) and not is_position(k, a):
                    new[k] = "_xlfn.TOCOL(%s)" % a.strip(); touched = True
            if touched:
                rows.append((sn, addr, "tocol", raw, rebuild(new), []))
            # cellref
            new, touched, helpers = list(args), False, []
            for k, a in enumerate(args):
                if is_position(k, a):
                    continue
                ms, mn = STR_LIT.match(a), NUM_LIT.match(a)
                if (ms or mn) and len(helpers) < len(HELPER_COLS):
                    helpers.append(ms.group(1).replace('""', '"') if ms else float(a))
                    new[k] = f"{SHEET}!$%s$@ROW@" % HELPER_COLS[len(helpers) - 1]; touched = True
            if touched:
                rows.append((sn, addr, "cellref", raw, rebuild(new), helpers))
            # textnum
            new, touched = list(args), False
            for k, a in enumerate(args):
                if is_position(k, a) or (fn, k + 1) in DATA_ARGS or (sn, addr) in TEXTNUM_SKIP_CASES:
                    continue
                if NUM_LIT.match(a):
                    new[k] = '"' + a.strip() + '"'; touched = True
                elif CELL_REF.match(a) and a.strip().replace("$", "") in numeric_cells:
                    new[k] = "(" + a.strip() + '&"")'; touched = True
            if touched:
                rows.append((sn, addr, "textnum", raw, rebuild(new), []))
            # array
            new, touched = list(args), False
            for k, a in enumerate(args):
                if RANGE_ARG.match(a) and not is_position(k, a) and not range_has_blank(a):
                    r = a.strip(); new[k] = f"_xlfn.CHOOSEROWS({r},_xlfn.SEQUENCE(ROWS({r})))"; touched = True
            if touched:
                rows.append((sn, addr, "array", raw, rebuild(new), []))
            # blank
            total, nopt = sig.get(fn, (len(args), 0))
            new, touched = list(args), False
            for k, a in enumerate(args):
                if a.strip() == "" and k >= total - nopt and (fn, k + 1) not in BLANK_SKIP:
                    new[k] = BLANK_REF; touched = True
            while nopt > 0 and total - nopt <= len(new) < total:
                if (fn, len(new) + 1) in BLANK_SKIP:
                    new.append("")
                else:
                    new.append(BLANK_REF)
                touched = True
            while new and new[-1] == "":
                new.pop()
            if touched:
                rows.append((sn, addr, "blank", raw, rebuild(new), []))
    existing_cells, donor_cells = [], []
    if SHEET in order:
        x = z.read(sheet_part(z, wbx, rels, SHEET)).decode()
        existing_cells = [a for a in re.findall(r'<c r="([A-Z]+\d+)"', x) if int(re.search(r"\d+", a).group()) >= 4]
    else:
        x = z.read(sheet_part(z, wbx, rels, TEMPLATE_SHEET)).decode()
        donor_cells = [a for a in re.findall(r'<c r="([A-Z]+\d+)"', x) if int(re.search(r"\d+", a).group()) >= 4]
    lt = z.read(sheet_part(z, wbx, rels, "lambda_tests")).decode()
    return order, lam_sheets, rows, existing_cells, donor_cells, lt


def match_formula(orig_q: str, var_q: str) -> str:
    """PASS when original and variant agree: same shape and every element equal (errors by ERROR.TYPE).
    A scalar error has no ROWS(), so shapes are read through IFERROR -> "err" and compared first."""
    a, b, sa, sb = "_xlpm.a", "_xlpm.b", "_xlpm.sa", "_xlpm.sb"
    shape = lambda v: 'IFERROR(ROWS(%s)&"x"&COLUMNS(%s),"err")' % (v, v)
    return ("_xlfn.LET(%s,%s,%s,%s,%s,%s,%s,%s,"
            "IF(%s<>%s,\"FAIL\",IF(%s=\"err\",IF(ERROR.TYPE(%s)=ERROR.TYPE(%s),\"PASS\",\"FAIL\"),"
            "IF(AND(IFERROR(%s=%s,IFERROR(ERROR.TYPE(%s)=ERROR.TYPE(%s),FALSE))),\"PASS\",\"FAIL\"))))"
            % (a, orig_q, b, var_q, sa, shape(a), sb, shape(b), sa, sb, sa, a, b, a, b, a, b))


def main(src: str, out: str) -> int:
    order, lam_sheets, rows, existing_cells, donor_cells, lt = harvest(src)
    new_sheet = SHEET not in order
    edits = []

    def put(cell, value=None, formula=None, array_ref=None):
        d = {"sheet": SHEET, "cell": cell}
        if formula is not None:
            d["formula"] = formula
            d["array_ref"] = array_ref
        else:
            d["value"] = value
        edits.append(d)

    for a in (donor_cells if new_sheet else existing_cells):
        edits.append({"sheet": SHEET, "cell": a})
    if new_sheet:
        put("B1", "shape tests  (generated - do not edit by hand)")
        put("B2", "py tools/build_shape_tests.py  <src.xlsm> <out.xlsm>   re-generates this sheet from the other test sheets")
        put("I1", "Each row re-runs one existing test case with its arguments transformed in a way that must not change the answer (1x1 array, value in a cell, number as text, computed array instead of range, blank cell for an omitted optional) and PASSes when the result matches in shape and value.")
    fam_counts = {}
    for r in rows:
        fam_counts[r[2]] = fam_counts.get(r[2], 0) + 1
    put("A4", "shape_tests - families: tocol (1x1 array), cellref (literal moved to a cell), textnum (number as text), array (computed array for a range), blank (empty cell for an omitted optional). Each must equal the plain call. Excluded on purpose: position-only range args, data lookup values, exlookup's if_not_found. "
        + "  ".join(f"{k}={v}" for k, v in sorted(fam_counts.items())))
    put("F4", formula='COUNTIF(F6:F5000,"PASS")&" / %d PASS"' % len(rows))
    for c, v in zip("ABCDEF", ["#", "Lambda", "Source case", "Family", "Variant", "Match"]):
        put(f"{c}5", v)
    put("H5", "helper values (cellref)")
    for n, (sn, addr, fam, raw, variant, helpers) in enumerate(rows, 1):
        r = 5 + n
        variant = variant.replace("@ROW@", str(r))
        put(f"A{r}", n)
        put(f"B{r}", sn)
        put(f"C{r}", f"{sn}!{addr}")
        put(f"D{r}", fam)
        put(f"E{r}", variant.replace("_xlfn.", "").replace("_xlpm.", ""))
        for k, h in enumerate(helpers):
            put(f"{HELPER_COLS[k]}{r}", h)
        put(f"F{r}", formula=match_formula(qualify(raw, sn), qualify(variant, sn)), array_ref=f"F{r}")

    dups, new_order = None, None
    if new_sheet:
        dups = [(TEMPLATE_SHEET, SHEET)]
        names = sorted(lam_sheets + [SHEET])
        new_order = list(order)
        prev = names[names.index(SHEET) - 1]
        new_order.insert(new_order.index(prev) + 1, SHEET)

        def style_of(addr):
            m = re.search(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>)' % addr, lt)
            s = re.search(r' s="(\d+)"', m.group(0)) if m else None
            return int(s.group(1)) if s else None
        old_total = 5 + len(lam_sheets)
        normal = {c: style_of(f"{c}{old_total - 1}") for c in "ABCDE"}
        total = {c: style_of(f"{c}{old_total}") for c in "ABCDE"}
        first = 5 + names.index(SHEET)
        for i, nm in enumerate(names):
            rr = 5 + i
            if rr < first:
                continue
            edits += [
                {"sheet": "lambda_tests", "cell": f"A{rr}", "formula": f"HYPERLINK(\"#'{nm}'!A1\",\"{nm}\")", "style": normal["A"]},
                {"sheet": "lambda_tests", "cell": f"B{rr}", "formula": f'COUNTIF({nm}!$A:$BZ,"PASS")', "style": normal["B"]},
                {"sheet": "lambda_tests", "cell": f"C{rr}", "formula": f'COUNTIF({nm}!$A:$BZ,"FAIL")', "style": normal["C"]},
                {"sheet": "lambda_tests", "cell": f"D{rr}", "formula": f"B{rr}+C{rr}", "style": normal["D"]},
                {"sheet": "lambda_tests", "cell": f"E{rr}", "formula": f'IF(C{rr}>0,"FAIL ("&C{rr}&")",IF(D{rr}=0,"-","all pass"))', "style": normal["E"]},
            ]
        nt = 5 + len(names)
        edits += [
            {"sheet": "lambda_tests", "cell": f"A{nt}", "value": "TOTAL", "style": total["A"]},
            {"sheet": "lambda_tests", "cell": f"B{nt}", "formula": f"SUM(B5:B{nt - 1})", "style": total["B"]},
            {"sheet": "lambda_tests", "cell": f"C{nt}", "formula": f"SUM(C5:C{nt - 1})", "style": total["C"]},
            {"sheet": "lambda_tests", "cell": f"D{nt}", "formula": f"SUM(D5:D{nt - 1})", "style": total["D"]},
            {"sheet": "lambda_tests", "cell": f"E{nt}", "formula": f'IF(C{nt}>0,C{nt}&" FAILING","ALL PASS")', "style": total["E"]},
        ]
    written = {(e["sheet"], e["cell"]) for e in edits if "value" in e or "formula" in e}
    edits = [e for e in edits if ("value" in e or "formula" in e) or (e["sheet"], e["cell"]) not in written]
    apply_edits(src, out, edits, duplicates=dups, sheet_order=new_order)
    res = validate(src, out, edits, duplicates=dups, sheet_order=new_order)
    print(f"shape_tests: {len(rows)} rows from {len(lam_sheets)} sheets {fam_counts}; new sheet: {new_sheet}")
    print("validate:", json.dumps({k: res[k] for k in ("ok", "edits_verified", "edits_failed", "unexpected_changes")}, default=str)[:400])
    return 0 if res.get("ok") else 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2]))
