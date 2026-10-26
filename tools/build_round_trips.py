"""Build (or rebuild) the `round_trips` test sheet: chains of library calls that must be identities.

Each case chains two or more lambdas the way a competition formula does (address <-> numeric
address, value -> address -> value, save -> load, move/expand there and back, ...) and checks
the chain comes back to where it started, so the HANDOFF between lambdas is tested, not just
each lambda alone. Edit the `cases` list below, then:

    py tools/build_round_trips.py "<src.xlsm>" "<out.xlsm>"

On first use it creates the sheet (from the count_substring layout), its roll-up row and tab
position; afterwards it rewrites the rows. Re-run tools/build_shape_tests.py afterwards - it
harvests these chains into its shape families as well.
"""
import sys, re, zipfile, json
sys.path.insert(0, r"C:\MEWC Dev Cowork\mewc-lambdas\tools")
from xlsm_edit import apply_edits, validate
SRC = sys.argv[1]; OUT = sys.argv[2]
SHEET = "round_trips"; DONOR = "count_substring"
z = zipfile.ZipFile(SRC)
wbx = z.read("xl/workbook.xml").decode(); rels = z.read("xl/_rels/workbook.xml.rels").decode()
def part(name):
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wbx).group(1)
    m = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, rels).group(0)
    return "xl/" + re.search(r'Target="([^"]+)"', m).group(1).replace("/xl/", "").lstrip("/")
order = re.findall(r'<sheet [^>]*name="([^"]+)"', wbx)
new_sheet = SHEET not in order
lam_names = [n for n in order if n not in ("Prep", "Lamb", "lambda_tests", "vba_tests") and not n.startswith("zz_")]
new_names = sorted(lam_names + [SHEET])
new_order = list(order); prev = new_names[new_names.index(SHEET) - 1]; new_order.insert(new_order.index(prev) + 1, SHEET)

E = []
def put(cell, value=None, formula=None, array_ref=None):
    d = {"sheet": SHEET, "cell": cell}
    if formula is not None: d["formula"] = formula; d["array_ref"] = array_ref
    else: d["value"] = value
    E.append(d)
# clear rows >= 4 of the donor (first run) or of the existing sheet (rebuild)
x = z.read(part(DONOR if new_sheet else SHEET)).decode()
for a in re.findall(r'<c r="([A-Z]+\d+)"', x):
    if int(re.search(r"\d+", a).group()) >= 4: E.append({"sheet": SHEET, "cell": a})
if new_sheet: put("B1", "round trips  -  chains of library calls that must come back to where they started")
if new_sheet: put("B2", "hand-written cases (tools/build_round_trips.py; tools/build_shape_tests.py also re-runs them in its shape families)")
if new_sheet: put("I1", "Each case chains two or more lambdas the way a competition formula does and checks the chain is an identity, so the HANDOFF between lambdas (numeric vs A1 addresses, single address vs range string, ';'-joined lists, text-form numbers) is tested, not just each lambda alone. Fixture grid: D90:F92 (20 appears twice).")
put("A4", "round_trips - identities across lambda handoffs. Fixture D90:F92 = 10,20,30 / 40,50,20 / 70,80,90 (20 appears in E90 and F91).")
for c, v in zip("A B D J L N".split(), ["#", "Case", "Chain", "Expected", "Actual", "Match"]): put(f"{c}5", v)
# fixture
grid = [[10, 20, 30], [40, 50, 20], [70, 80, 90]]
for i, row in enumerate(grid):
    for j, v in enumerate(row): put(f"{'DEF'[j]}{90 + i}", v)
put("D89", "fixture grid (do not move)")

T = "_xlfn."
cases = [
 # (desc, chain text, expected, actual formula, kind)  kind: "eq" scalar compare | "bool" expects TRUE | ("spill", rows) column spill
 ("numeric -> A1 -> numeric", "num_ad(str_ad(3000004))", 3000004, "num_ad(str_ad(3000004))", "eq"),
 ("A1 -> numeric -> A1", 'str_ad(num_ad("AM17"))', "AM17", 'str_ad(num_ad("AM17"))', "eq"),
 ("messy A1 ($, lower case) normalises through the pair", 'str_ad(num_ad("$am$17"))', "AM17", 'str_ad(num_ad("$am$17"))', "eq"),
 ("range: numeric -> A1 -> numeric (2x2)", 'num_ad(str_ad(num_ad("B2:C3"))) = num_ad("B2:C3")', True, 'AND(num_ad(str_ad(num_ad("B2:C3")))=num_ad("B2:C3"))', "bool"),
 ("value -> address -> value (unique value)", "ad_to_cont(cont_to_ad(50,grid),grid)", 50, "ad_to_cont(cont_to_ad(50,D90:F92),D90:F92)", "eq"),
 ("value -> A1 address -> value (str=1)", "ad_to_cont(cont_to_ad(80,grid,1),grid)", 80, "ad_to_cont(cont_to_ad(80,D90:F92,1),D90:F92)", "eq"),
 ("duplicate value -> ';'-joined addresses -> split (TEXT pieces) -> values", 'ad_to_cont(TOCOL(TEXTSPLIT(cont_to_ad(20,grid),";")),grid)', [20, 20], f'ad_to_cont({T}TOCOL({T}TEXTSPLIT(cont_to_ad(20,D90:F92),";")),D90:F92)', ("spill", 2)),
 ("split text addresses -> num_ad gives the numbers", 'num_ad(TEXTSPLIT(cont_to_ad(20,grid),";")) = {90000005,91000006}', True, f'AND(num_ad({T}TEXTSPLIT(cont_to_ad(20,D90:F92),";"))={{90000005,91000006}})', "bool"),
 ("save_array -> load_array (default delimiters)", "load_array(save_array(grid)) = grid", True, "AND(load_array(save_array(D90:F92))=D90:F92)", "bool"),
 ("save_array -> load_array (custom delimiters)", 'load_array(save_array(grid,",","/"),",","/") = grid', True, 'AND(load_array(save_array(D90:F92,",","/"),",","/")=D90:F92)', "bool"),
 ("move right 3 then left 3 (range string)", 'move_rng(move_rng("B2:C10","→",3),"←",3)', "B2:C10", 'move_rng(move_rng("B2:C10","→",3),"←",3)', "eq"),
 ("move ↘ then ↖ (single address)", 'move_rng(move_rng("B2","↘"),"↖")', "B2", 'move_rng(move_rng("B2","↘"),"↖")', "eq"),
 ("grow 2,2 then shrink -2,-2 back to one cell", 'expand_rng(expand_rng("B3",2,2),-2,-2)', "B3", 'expand_rng(expand_rng("B3",2,2),-2,-2)', "eq"),
 ("grow all sides then shrink all sides", 'expand_rng(expand_rng("B2:D6",1,1,1,1),-1,-1,-1,-1)', "B2:D6", 'expand_rng(expand_rng("B2:D6",1,1,1,1),-1,-1,-1,-1)', "eq"),
 ("rotate 3 clockwise then 3 counter-clockwise", 'rotate_arrow(rotate_arrow("→",3),3,0)', "→", 'rotate_arrow(rotate_arrow("→",3),3,0)', "eq"),
 ("rotate a full circle", 'rotate_arrow("↑",8)', "↑", 'rotate_arrow("↑",8)', "eq"),
 ("arrow_move ↘ then ↖, back through str_ad", 'str_ad(arrow_move(arrow_move("C3","↘"),"↖"))', "C3", 'str_ad(arrow_move(arrow_move("C3","↘"),"↖"))', "eq"),
 ("rotate_arrow output drives arrow_move", 'arrow_move("C3",rotate_arrow("↑",2)) = arrow_move("C3","→")', True, 'arrow_move("C3",rotate_arrow("↑",2))=arrow_move("C3","→")', "bool"),
 ("expand_rng single-address output consumed by ad_to_cont", 'ad_to_cont(expand_rng("E91"),grid)', 50, 'ad_to_cont(expand_rng("E91"),D90:F92)', "eq"),
 ("expand_rng block output consumed by ad_to_cont", 'ad_to_cont(expand_rng("D90",1,1),grid) = D90:E91', True, 'AND(ad_to_cont(expand_rng("D90",1,1),D90:F92)=D90:E91)', "bool"),
 ("move_rng block output consumed by ad_to_cont", 'ad_to_cont(move_rng("D90:E91","↘"),grid) = E91:F92', True, 'AND(ad_to_cont(move_rng("D90:E91","↘"),D90:F92)=E91:F92)', "bool"),
 ("neighbors_list output consumed by ad_to_cont (sum of E91's 8 neighbours)", 'SUM(ad_to_cont(neighbors_list("E91",1,grid),grid))', 360, 'SUM(ad_to_cont(neighbors_list("E91",1,D90:F92),D90:F92))', "eq"),
 ("arrow_move output consumed by dist", 'dist(arrow_move("A1","↘",3),"A1")', 3, 'dist(arrow_move("A1","↘",3),"A1")', "eq"),
 ("row/col numbers -> rc_to_ad -> str_ad", 'str_ad(rc_to_ad(ad_row_n("AM17"),ad_col_n("AM17")))', "AM17", 'str_ad(rc_to_ad(ad_row_n("AM17"),ad_col_n("AM17")))', "eq"),
 ("ad_to_rc -> rc_to_ad", 'str_ad(rc_to_ad(INDEX(ad_to_rc("D4"),1),INDEX(ad_to_rc("D4"),2)))', "D4", 'str_ad(rc_to_ad(INDEX(ad_to_rc("D4"),1),INDEX(ad_to_rc("D4"),2)))', "eq"),
 ("column letters -> number -> letters", 'ad_col_let(ad_col_n("BDS1"))', "BDS", 'ad_col_let(ad_col_n("BDS1"))', "eq"),
 ("xtextjoin string mode -> array mode", 'xtextjoin(xtextjoin({"a";"b";"c"}),";",2) = {"a";"b";"c"}', True, 'AND(xtextjoin(xtextjoin({"a";"b";"c"}),";",2)={"a";"b";"c"})', "bool"),
 ("str_to_arr -> xtextjoin -> str_to_arr", 'xtextjoin(str_to_arr("abc"),"-")', "a-b-c", 'xtextjoin(str_to_arr("abc"),"-")', "eq"),
 ("str_rev twice", 'str_rev(str_rev("hello"))', "hello", 'str_rev(str_rev("hello"))', "eq"),
 ("xdec2bin -> BIN2DEC", "BIN2DEC(xdec2bin(37))", 37, "BIN2DEC(xdec2bin(37))", "eq"),
 ("offset_letter +3 then -3", 'offset_letter(offset_letter("abc",3),-3)', "abc", 'offset_letter(offset_letter("abc",3),-3)', "eq"),
 ("shift_arr down then up", 'shift_arr(shift_arr({1;2;3},"↓"),"↑") = {1;2;3}', True, 'AND(shift_arr(shift_arr({1;2;3},"↓"),"↑")={1;2;3})', "bool"),
 ("transform_array ↻ then ↺", 'transform_array(transform_array({1,2;3,4},"↻"),"↺") = {1,2;3,4}', True, 'AND(transform_array(transform_array({1,2;3,4},"↻"),"↺")={1,2;3,4})', "bool"),
 ("save -> load -> set", 'set(load_array(save_array({3;1;2;3}))) = {1;2;3}', True, 'AND(set(load_array(save_array({3;1;2;3})))={1;2;3})', "bool"),
 ("two arrow_moves, then dist between them", 'dist(arrow_move("D4","↑",2),arrow_move("D4","→",1))', 2, 'dist(arrow_move("D4","↑",2),arrow_move("D4","→",1))', "eq"),
]
r = 6
for n, (desc, chain, expected, formula, kind) in enumerate(cases, 1):
    put(f"A{r}", n); put(f"B{r}", desc); put(f"D{r}", chain)
    if isinstance(kind, tuple):
        rows = kind[1]
        for i, v in enumerate(expected): put(f"J{r+i}", v)
        put(f"L{r}", formula=formula, array_ref=f"L{r}:L{r+rows-1}")
        put(f"N{r}", formula=f'IFERROR(IF(AND(J{r}:J{r+rows-1}=L{r}:L{r+rows-1}),"PASS","FAIL"),"FAIL")', array_ref=f"N{r}")
        r += rows + 1
    else:
        put(f"J{r}", expected)
        put(f"L{r}", formula=formula, array_ref=f"L{r}")
        put(f"N{r}", formula=f'IFERROR(IF(L{r}=J{r},"PASS","FAIL"),"FAIL")')
        r += 2
assert r <= 89, r
put("N4", formula=f'COUNTIF(N6:N88,"PASS")&" / {len(cases)} PASS"')
# roll-up (first run only)
lt = z.read(part("lambda_tests")).decode()
if not new_sheet:
    new_names = lam_names; nt = 5 + len(lam_names); first = 10**9
def style_of(addr):
    m = re.search(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>)' % addr, lt); s = re.search(r' s="(\d+)"', m.group(0)) if m else None
    return int(s.group(1)) if s else None
old_total = 5 + len(lam_names); normal = {c: style_of(f"{c}{old_total-1}") for c in "ABCDE"}; total = {c: style_of(f"{c}{old_total}") for c in "ABCDE"}
first = 5 + new_names.index(SHEET)
for i, nm in enumerate(new_names):
    rr = 5 + i
    if rr < first: continue
    E += [{"sheet": "lambda_tests", "cell": f"A{rr}", "formula": f"HYPERLINK(\"#'{nm}'!A1\",\"{nm}\")", "style": normal["A"]},
          {"sheet": "lambda_tests", "cell": f"B{rr}", "formula": f'COUNTIF({nm}!$A:$BZ,"PASS")', "style": normal["B"]},
          {"sheet": "lambda_tests", "cell": f"C{rr}", "formula": f'COUNTIF({nm}!$A:$BZ,"FAIL")', "style": normal["C"]},
          {"sheet": "lambda_tests", "cell": f"D{rr}", "formula": f"B{rr}+C{rr}", "style": normal["D"]},
          {"sheet": "lambda_tests", "cell": f"E{rr}", "formula": f'IF(C{rr}>0,"FAIL ("&C{rr}&")",IF(D{rr}=0,"-","all pass"))', "style": normal["E"]}]
nt = 5 + len(new_names)
if new_sheet: E += [{"sheet": "lambda_tests", "cell": f"A{nt}", "value": "TOTAL", "style": total["A"]},
      {"sheet": "lambda_tests", "cell": f"B{nt}", "formula": f"SUM(B5:B{nt-1})", "style": total["B"]},
      {"sheet": "lambda_tests", "cell": f"C{nt}", "formula": f"SUM(C5:C{nt-1})", "style": total["C"]},
      {"sheet": "lambda_tests", "cell": f"D{nt}", "formula": f"SUM(D5:D{nt-1})", "style": total["D"]},
      {"sheet": "lambda_tests", "cell": f"E{nt}", "formula": f'IF(C{nt}>0,C{nt}&" FAILING","ALL PASS")', "style": total["E"]}]
written = {(e["sheet"], e["cell"]) for e in E if "value" in e or "formula" in e}
E = [e for e in E if ("value" in e or "formula" in e) or (e["sheet"], e["cell"]) not in written]
dups = [(DONOR, SHEET)] if new_sheet else None
if not new_sheet: new_order = None
apply_edits(SRC, OUT, E, duplicates=dups, sheet_order=new_order)
res = validate(SRC, OUT, E, duplicates=dups, sheet_order=new_order)
print(f"round_trips: {len(cases)} cases, last row {r}; TOTAL -> row {nt}")
print("validate:", json.dumps({k: res[k] for k in ("ok", "edits_verified", "edits_failed", "unexpected_changes")}, default=str)[:300])
