#!/usr/bin/env python3
"""
lambda_check.py — sanity-checker for MEWC custom LAMBDAs.

Scans a lambda definition and flags the things that commonly break lambdas or
violate this project's authoring rules:

  * variable / parameter names that are valid cell references (x2, num1, c1 ...)
  * names that collide with a built-in function (N, T, SUM, ...) or with an
    existing workbook defined name / lambda (set, dist, arrow, ...)
  * inline comment markers ( //  /*  */ )
  * optional arguments that are ISOMITTED-tested but not wrapped in [brackets]
  * required arguments placed after an optional one
  * definitions over the 2048-character Name Manager limit

USAGE
  python3 lambda_check.py mylambda.txt            # check one file
  python3 lambda_check.py a.txt b.txt             # check several
  echo '=LAMBDA(arr,[as_row], ...)' | python3 lambda_check.py -

The file should contain a single lambda definition (multi-line is fine).
Exit code is 0 if all files PASS, 1 if any file has an ERROR.
"""

import sys, os, re, glob

# ---------------------------------------------------------------------------
# workbook defined names (lambdas + ranges) — loaded from the template if found
# ---------------------------------------------------------------------------
LIB_LAMBDAS = set("""addr_lookup arrow arrow_move arrow_trip biggest clamp
count_substring ctotal cycles die_pips dist exists freq_table is_unique
knight_moves label_lookup lc_to_nc load_array nc_to_lc neighbors_count
neighbors_list offset_letter only_letters only_numbers rc_to_ref ref_to_rc rev
rotate_arrow route_cost save_array selections set set_d set_i set_u shift_arr
smallest str2arr str_rev textbetween transform_array update_arr xcountif
xdec2bin xmod xn xrank xsequence xtextsplit xtocol""".split())

def load_workbook_names(folder):
    names = set()
    for pat in ("MEWC Lambdas*.xlsm", "*Lambdas*VBA*.xlsm"):
        for p in glob.glob(os.path.join(folder, pat)):
            try:
                import zipfile, html
                z = zipfile.ZipFile(p)
                d = z.read("xl/workbook.xml").decode("utf-8")
                for m in re.finditer(r'<definedName name="([^"]+)"', d):
                    n = m.group(1)
                    if n.startswith("_xl") or n.startswith("IQ_"):
                        continue
                    names.add(n)
            except Exception:
                pass
    return names

# A broad set of Excel function names to guard against shadowing.
BUILTINS = set("""N T C R ABS AND AVERAGE BASE BYCOL BYROW CEILING CHAR CHOOSE
CHOOSECOLS CHOOSEROWS CODE COLUMN COLUMNS CONCAT COUNT COUNTA COUNTIF DROP EXACT
EXPAND FILTER FIND HLOOKUP HSTACK IF IFERROR IFNA IFS INDEX INT ISBLANK ISERROR
ISLOGICAL ISNA ISNUMBER ISODD ISOMITTED ISTEXT LAMBDA LEFT LEN LET LOOKUP LOWER
MAKEARRAY MAP MATCH MAX MID MIN MOD NA OR PRODUCT QUOTIENT RANK REDUCE REGEXTEST
REGEXEXTRACT REGEXREPLACE RIGHT ROUND ROW ROWS SCAN SEQUENCE SIGN SORT SORTBY
SUБSTITUTE SUBSTITUTE SUM SUMPRODUCT SWITCH TAKE TEXT TEXTAFTER TEXTBEFORE TEXTJOIN
TEXTSPLIT TOCOL TOROW TRANSPOSE TRIM UNICHAR UNICODE UNIQUE UPPER VALUE VLOOKUP
VSTACK WRAPCOLS WRAPROWS XLOOKUP XMATCH""".split())

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def is_cell_ref(name):
    """True if name looks like a real A1 cell reference (breaks lambdas)."""
    m = re.fullmatch(r'([A-Za-z]{1,3})(\d+)', name)
    if not m:
        return False
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return col <= 16384 and 1 <= int(m.group(2)) <= 1048576

def split_top_args(s, open_idx):
    """Return list of top-level argument strings for the (...) starting at open_idx."""
    args, depth, cur, i, instr, q = [], 0, [], open_idx, False, ""
    while i < len(s):
        ch = s[i]
        if instr:
            cur.append(ch)
            if ch == q:
                instr = False
        elif ch in '"\'':
            instr, q = True, ch
            cur.append(ch)
        elif ch == '(':
            depth += 1
            if depth > 1:
                cur.append(ch)
        elif ch == ')':
            depth -= 1
            if depth == 0:
                args.append("".join(cur).strip())
                return args, i
            cur.append(ch)
        elif ch == ',' and depth == 1:
            args.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
        i += 1
    return args, i

def find_groups(s, keyword):
    """Yield (open_paren_index) for each `keyword(` occurrence (case-insensitive)."""
    for m in re.finditer(r'(?<![A-Za-z0-9_.])' + keyword + r'\s*\(', s, re.I):
        yield m.end() - 1

# ---------------------------------------------------------------------------
# core check
# ---------------------------------------------------------------------------
def check(formula, wbnames):
    errors, warns, infos = [], [], []
    raw = formula

    # 1. length
    n = len(formula)
    if n > 2048:
        errors.append(f"definition is {n} chars — over the 2048 limit (decompose it)")
    else:
        infos.append(f"length {n}/2048 chars")

    # 2. comment markers
    for mark in ("//", "/*", "*/"):
        if mark in formula:
            errors.append(f"contains comment marker '{mark}' — no comments allowed in the formula")

    # collect LAMBDA params (with bracket flag) and LET vars
    params = []          # list of (name, bracketed, order_index)
    let_vars = []
    # top-level lambda = first LAMBDA(
    lam_idxs = list(find_groups(formula, "LAMBDA"))
    top_params = []
    for gi, oi in enumerate(lam_idxs):
        args, _ = split_top_args(formula, oi)
        # all but last are parameters
        for a in args[:-1]:
            a = a.strip()
            br = a.startswith('[') and a.endswith(']')
            nm = a[1:-1].strip() if br else a
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', nm):
                params.append((nm, br))
                if gi == 0:
                    top_params.append((nm, br))
    for oi in find_groups(formula, "LET"):
        args, _ = split_top_args(formula, oi)
        for a in args[:-1][::2]:
            a = a.strip()
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', a):
                let_vars.append(a)

    allnames = [(nm, "param") for nm, _ in params] + [(v, "LET var") for v in let_vars]

    # 3. name checks
    seen = set()
    for nm, kind in allnames:
        key = (nm, kind)
        if key in seen:
            continue
        seen.add(key)
        if is_cell_ref(nm):
            errors.append(f"{kind} '{nm}' is a valid cell reference — will break the lambda")
        if nm.upper() in BUILTINS:
            warns.append(f"{kind} '{nm}' shadows built-in function {nm.upper()}()")
        if nm in wbnames or nm in LIB_LAMBDAS:
            warns.append(f"{kind} '{nm}' collides with an existing workbook name/lambda")
        if re.fullmatch(r'[A-Za-z]', nm):
            infos.append(f"{kind} '{nm}' is a single letter — legal but discouraged for clarity")

    # 4. optional-argument bracket check (top-level signature)
    top_names = {nm for nm, _ in top_params}
    bracketed = {nm for nm, br in top_params if br}
    isomitted = set(re.findall(r'ISOMITTED\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)', formula, re.I))
    for nm in isomitted:
        if nm in top_names and nm not in bracketed:
            errors.append(f"'{nm}' is ISOMITTED-tested but not wrapped in [brackets] in the signature")
    # required-after-optional ordering
    seen_opt = None
    for nm, br in top_params:
        if br:
            seen_opt = nm
        elif seen_opt:
            errors.append(f"required parameter '{nm}' comes after optional '[{seen_opt}]' — put optionals last")
            break

    return errors, warns, infos

# ---------------------------------------------------------------------------
def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 0
    folder = os.path.dirname(os.path.abspath(args[0])) if args[0] != "-" else "."
    wbnames = load_workbook_names(folder)

    any_error = False
    for path in args:
        if path == "-":
            text, label = sys.stdin.read(), "<stdin>"
        else:
            text, label = open(path, encoding="utf-8").read(), os.path.basename(path)
        if "=== CODE ===" in text:
            _seg = text.split("=== CODE ===", 1)[1].lstrip("\n")
            _lines = []
            for _ln in _seg.splitlines():
                _s = _ln.strip()
                if _s.startswith("=== ") and _s.endswith(" ==="):
                    break
                _lines.append(_ln)
            formula = "\n".join(_lines).strip()
        else:
            formula = text.strip()
        errors, warns, infos = check(formula, wbnames)
        status = "FAIL" if errors else ("WARN" if warns else "PASS")
        print(f"\n=== {label} : {status} ===")
        for e in errors: print(f"  ERROR  {e}")
        for w in warns:  print(f"  WARN   {w}")
        for i in infos:  print(f"  info   {i}")
        if errors:
            any_error = True
    return 1 if any_error else 0

if __name__ == "__main__":
    sys.exit(main())
