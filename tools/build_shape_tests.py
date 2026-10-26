#!/usr/bin/env python3
"""Build (or rebuild) the `shape_tests` sheet in the test workbook.

What it tests
-------------
Every lambda test case on the per-lambda sheets is re-run with each of its
single-value arguments wrapped in TOCOL(...) - i.e. passed as a 1x1 ARRAY instead
of a plain value - and the result is compared with the original. A 1x1 array is
what a caller gets from TOCOL(x), an array literal {x}, or a one-element
FILTER/INDEX/TEXTSPLIT result, and Excel does not unbox it: SEQUENCE(n),
MAKEARRAY(n, ...), REGEXEXTRACT(x, ..., 1) and IF(flag, a, b) all *lift* over it
and silently return a single element or a doubled result. (z_a1n once read "AM17"
as column Z this way.) The library guards against it with z_scalar / CONCAT at the
top of each scalar argument; this sheet is the regression net for that guard.

One row per harvested case: the row's Match cell recomputes BOTH the original
formula and the all-arguments-wrapped variant (with cell references qualified to
the source sheet) and PASSes when they agree in shape and in every element
(errors compare by ERROR.TYPE).

Arguments that are NOT wrapped (wrapping would legitimately change the meaning):
  * range-reference arguments whose POSITION matters - ad_to_cont's arr,
    arrow_move's board_rng, neighbors_list's boundary_rng, and a cell reference
    passed as move_rng's ad_or_rng (TOCOL turns a reference into data);
  * anything that is already a range (A1:B2), an array literal, or an expression.

Usage
-----
    py tools/build_shape_tests.py "<src.xlsm>" "<out.xlsm>"

Re-run it whenever test cases are added or changed; it rewrites the sheet's rows
(creating the sheet, its roll-up row and tab position on first use). Then open the
workbook in Excel, run sync_all / lambda_update as usual, and check the tally.
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsm_edit import apply_edits, validate  # noqa: E402

SHEET = "shape_tests"
TEMPLATE_SHEET = "count_substring"          # layout donor for a fresh sheet
NON_LAMBDA_SHEETS = {"Prep", "Lamb", "lambda_tests", "vba_tests", SHEET}
# (lambda, 1-based argument index) whose argument must stay a reference
EXCLUDE = {("ad_to_cont", 2), ("arrow_move", 4), ("neighbors_list", 3)}
REF_ONLY_FIRST_ARG = {"move_rng"}            # skip arg 1 only when it is a bare cell reference

WRAPPABLE = re.compile(r'^\s*(\$?[A-Z]{1,3}\$?\d+|"(?:[^"]|"")*"|-?\d+(?:\.\d+)?)\s*$')
CELL_REF = re.compile(r'^\s*\$?[A-Z]{1,3}\$?\d+\s*$')
REF_TOKEN = re.compile(r"(?<![A-Za-z_!'.$])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Za-z_(\d])")


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
    """(start, end) of the first top-level `name(...)` call whose name is a whole word."""
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
    out, inq, i = [], False, 0
    buf = ""
    for ch in formula:
        if ch == '"':
            if not inq:
                out.append(REF_TOKEN.sub(lambda m: f"'{sheet}'!{m.group(1)}", buf))
                buf = ""
            else:
                out.append(buf)
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


def harvest(src: str):
    z = zipfile.ZipFile(src)
    wbx = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    order = re.findall(r'<sheet [^>]*name="([^"]+)"', wbx)
    lam_sheets = [n for n in order if n not in NON_LAMBDA_SHEETS and not n.startswith("zz_")]
    cases = []
    for sn in lam_sheets:
        x = z.read(sheet_part(z, wbx, rels, sn)).decode()
        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)>(.*?)</c>', x, re.S):
            addr, _attrs, inner = m.groups()
            fm = re.search(r"<f[^>]*>(.*?)</f>", inner, re.S)
            if not fm:
                continue
            raw = unescape(fm.group(1))
            if addr in ("B1", "B2", "I1", "S1") or raw.startswith(("COUNTIF(", "IF(", "HYPERLINK(")):
                continue
            span = find_call(raw, sn)
            if not span:
                continue
            i, j = span
            args = split_args(raw[i + len(sn) + 1:j - 1])
            wrap = []
            for k, a in enumerate(args):
                if not WRAPPABLE.match(a):
                    continue
                if (sn, k + 1) in EXCLUDE:
                    continue
                if sn in REF_ONLY_FIRST_ARG and k == 0 and CELL_REF.match(a):
                    continue
                wrap.append(k)
            if not wrap:
                continue
            new_args = [("_xlfn.TOCOL(%s)" % a.strip()) if k in wrap else a for k, a in enumerate(args)]
            variant = raw[:i] + sn + "(" + ",".join(new_args) + ")" + raw[j:]
            cases.append({"sheet": sn, "cell": addr, "original": raw, "variant": variant})
    existing_cells = []
    if SHEET in order:
        x = z.read(sheet_part(z, wbx, rels, SHEET)).decode()
        existing_cells = [a for a in re.findall(r'<c r="([A-Z]+\d+)"', x) if int(re.search(r"\d+", a).group()) >= 4]
    donor_cells = []
    if SHEET not in order:
        x = z.read(sheet_part(z, wbx, rels, TEMPLATE_SHEET)).decode()
        donor_cells = [a for a in re.findall(r'<c r="([A-Z]+\d+)"', x) if int(re.search(r"\d+", a).group()) >= 4]
    lt = z.read(sheet_part(z, wbx, rels, "lambda_tests")).decode()
    return order, lam_sheets, cases, existing_cells, donor_cells, lt


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
    order, lam_sheets, cases, existing_cells, donor_cells, lt = harvest(src)
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
        put("B1", "1x1-array shape tests  (generated - do not edit by hand)")
        put("B2", "py tools/build_shape_tests.py  <src.xlsm> <out.xlsm>   re-generates this sheet from the other test sheets")
        put("I1", "Each row re-runs one existing test case with every single-value argument wrapped in TOCOL() - a 1x1 array - and PASSes when the result matches the plain call in shape and value. Guards the z_scalar / CONCAT unboxing in the library.")
    put("A4", "shape_tests - every single-value argument passed as a 1x1 array (TOCOL) must give the same answer as the plain value. Excluded: reference arguments whose position matters (ad_to_cont arr, arrow_move board_rng, neighbors_list boundary_rng, a cell reference as move_rng's ad_or_rng).")
    put("F4", formula='COUNTIF(F6:F2000,"PASS")&" / %d PASS"' % len(cases))
    for c, v in zip("ABCDF", ["#", "Lambda", "Source case", "Variant (TOCOL-wrapped arguments)", "Match"]):
        put(f"{c}5", v)
    for n, cs in enumerate(cases, 1):
        r = 5 + n
        put(f"A{r}", n)
        put(f"B{r}", cs["sheet"])
        put(f"C{r}", f"{cs['sheet']}!{cs['cell']}")
        put(f"D{r}", cs["variant"].replace("_xlfn.", "").replace("_xlpm.", ""))
        put(f"F{r}", formula=match_formula(qualify(cs["original"], cs["sheet"]), qualify(cs["variant"], cs["sheet"])), array_ref=f"F{r}")

    dups, new_order = None, None
    if new_sheet:
        dups = [(TEMPLATE_SHEET, SHEET)]
        names = sorted(lam_sheets + [SHEET])
        new_order = list(order)
        prev = names[names.index(SHEET) - 1]
        new_order.insert(new_order.index(prev) + 1, SHEET)
        # roll-up: rewrite from the inserted position; TOTAL shifts down one row
        def style_of(addr):
            m = re.search(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>)' % addr, lt)
            s = re.search(r' s="(\d+)"', m.group(0)) if m else None
            return int(s.group(1)) if s else None
        old_total = 5 + len(lam_sheets)
        normal = {c: style_of(f"{c}{old_total - 1}") for c in "ABCDE"}
        total = {c: style_of(f"{c}{old_total}") for c in "ABCDE"}
        first = 5 + names.index(SHEET)
        for i, nm in enumerate(names):
            r = 5 + i
            if r < first:
                continue
            edits += [
                {"sheet": "lambda_tests", "cell": f"A{r}", "formula": f"HYPERLINK(\"#'{nm}'!A1\",\"{nm}\")", "style": normal["A"]},
                {"sheet": "lambda_tests", "cell": f"B{r}", "formula": f'COUNTIF({nm}!$A:$BZ,"PASS")', "style": normal["B"]},
                {"sheet": "lambda_tests", "cell": f"C{r}", "formula": f'COUNTIF({nm}!$A:$BZ,"FAIL")', "style": normal["C"]},
                {"sheet": "lambda_tests", "cell": f"D{r}", "formula": f"B{r}+C{r}", "style": normal["D"]},
                {"sheet": "lambda_tests", "cell": f"E{r}", "formula": f'IF(C{r}>0,"FAIL ("&C{r}&")",IF(D{r}=0,"-","all pass"))', "style": normal["E"]},
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
    print(f"shape_tests: {len(cases)} cases from {len(lam_sheets)} sheets; new sheet: {new_sheet}")
    print("validate:", json.dumps({k: res[k] for k in ("ok", "edits_verified", "edits_failed", "unexpected_changes")}, default=str)[:400])
    return 0 if res.get("ok") else 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2]))
