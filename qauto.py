"""Rebuild the 'Automation' sheet of Q-Automation.xlsx (formulas + cached values).

Only xl/worksheets/<Automation>.xml is rewritten; every other part of the workbook is copied untouched.
Automation!A1 = FILTER('Customer Wise'!A:M,'Customer Wise'!A:A<>"") ;  N:V are per-row lookups.
"""
import io, re, zipfile, datetime
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.utils.datetime import to_excel

ERR = "#N/A"


class NA:                                   # marker for an Excel #N/A result
    pass


def _ser(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return to_excel(v)
    return v


def _cells(v):
    return None if v == "" else _ser(v)


def _numtext(s):
    try:
        t = s.strip()
        if not t or t.lower() in ("nan", "inf", "-inf", "infinity", "+inf"):
            return None
        return float(t)
    except ValueError:
        return None


def ckey(v):
    """SUMIFS/COUNTIF style equality key (numeric-looking text == number, case-insensitive text)."""
    v = _ser(v)
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    if v is None:
        return ("n", 0.0)
    n = _numtext(str(v))
    return ("n", n) if n is not None else ("s", str(v).lower())


def xkey(v):
    """XLOOKUP exact-match key (type strict, case-insensitive text)."""
    v = _ser(v)
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    if v is None:
        return None
    return ("s", str(v).lower())


def _read(wb, name, maxcol):
    ws = wb[name]
    return [list(r) + [None] * (maxcol - len(r)) for r in ws.iter_rows(max_col=maxcol, values_only=True)]


def compute(src):
    """Return list of rows (A..V values) for the Automation sheet, header row first."""
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    cw = _read(wb, "Customer Wise", 13)
    sol = _read(wb, "Sales Order List", 13)
    pv = _read(wb, "Pivet Customer Dispatch", 12)
    fin = _read(wb, "Finalized", 9)
    wb.close()

    base = [[(0 if c is None else _ser(c)) for c in r[:13]]
            for r in cw if r[0] not in (None, "")]            # FILTER(A:M, A<>"")

    def first_map(rows, kc):
        d = {}
        for r in rows:
            k = xkey(r[kc])
            if k is not None:
                d.setdefault(k, r)
        return d
    sol_map, fin_map = first_map(sol, 0), first_map(fin, 3)

    pv_sum = {}
    for r in pv:                                               # SUMIFS(L, C=SO, F=item)
        q = r[11]
        if isinstance(q, (int, float)) and not isinstance(q, bool):
            k = (ckey(r[2]), ckey(r[5]))
            pv_sum[k] = pv_sum.get(k, 0) + q
    self_sum = {}                                              # SUMIFS(G:G, A:A, A, D:D, D)
    for r in base:
        q = r[6]
        if isinstance(q, (int, float)) and not isinstance(q, bool):
            k = (ckey(r[0]), ckey(r[3]))
            self_sum[k] = self_sum.get(k, 0) + q

    def look(m, key, idx):
        k = xkey(key)
        row = m.get(k) if k is not None else None
        if row is None:
            return NA
        v = row[idx]
        return 0 if v is None else _ser(v)

    out = []
    for i, r in enumerate(base):
        if i == 0:
            out.append(r + [None] * 9)                         # header handled separately
            continue
        k = (ckey(r[0]), ckey(r[3]))
        g = r[6] if isinstance(r[6], (int, float)) else 0
        p, s = pv_sum.get(k, 0), self_sum.get(k, 0)
        out.append(r + [look(sol_map, r[0], 11), look(sol_map, r[0], 12), p, g - p, s, g - s,
                        look(fin_map, r[2], 5), look(fin_map, r[2], 6), look(fin_map, r[2], 8)])
    return out


def _num(v):
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        v = int(v)
    return repr(v)


def _cell(ref, v, f=None):
    ftag = "<f>%s</f>" % escape(f) if f else ""
    if v is NA:
        return '<c r="%s" t="e">%s<v>#N/A</v></c>' % (ref, ftag)
    if isinstance(v, bool):
        return '<c r="%s" t="b">%s<v>%d</v></c>' % (ref, ftag, v)
    if isinstance(v, (int, float)):
        return '<c r="%s">%s<v>%s</v></c>' % (ref, ftag, _num(v))
    return '<c r="%s" t="str">%s<v>%s</v></c>' % (ref, ftag, escape(str(v)))


def _formulas(n):
    return {
        "N": "_xlfn.XLOOKUP($A{n},'Sales Order List'!$A:$A,'Sales Order List'!L:L)",
        "O": "_xlfn.XLOOKUP($A{n},'Sales Order List'!$A:$A,'Sales Order List'!M:M)",
        "P": "SUMIFS('Pivet Customer Dispatch'!$L:$L,'Pivet Customer Dispatch'!$C:$C,$A{n},'Pivet Customer Dispatch'!$F:$F,D{n})",
        "Q": "G{n}-P{n}",
        "R": "SUMIFS(G:G,A:A,A{n},D:D,D{n})",
        "S": "G{n}-R{n}",
        "T": "_xlfn.XLOOKUP($C{n},Finalized!$D:$D,Finalized!F:F)",
        "U": "_xlfn.XLOOKUP($C{n},Finalized!$D:$D,Finalized!G:G)",
        "V": "_xlfn.XLOOKUP($C{n},Finalized!$D:$D,Finalized!I:I)"}


def rebuild(src_bytes):
    rows = compute(io.BytesIO(src_bytes))
    n = len(rows)
    zin = zipfile.ZipFile(io.BytesIO(src_bytes))
    wbx = zin.read("xl/workbook.xml").decode("utf8")
    rel = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = re.search(r'<sheet name="Automation"[^>]*r:id="([^"]+)"', wbx)
    if not rid:
        raise ValueError("فائل میں 'Automation' نام کی شیٹ نہیں ملی")
    target = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"|<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"'
                       % (rid.group(1), rid.group(1)), rel)
    path = "xl/" + (target.group(1) or target.group(2)).lstrip("/").replace("xl/", "")
    sx = zin.read(path).decode("utf8")

    a, b = sx.index("<sheetData>"), sx.index("</sheetData>")
    row1 = re.search(r'<row r="1".*?</row>', sx[a:b], re.S).group(0)
    row1 = re.sub(r'(<f t="array" ref="A1:)M\d+(")', r"\g<1>M%d\2" % n, row1)
    forms = _formulas(0)
    body = []
    L = "ABCDEFGHIJKLM"
    for i, r in enumerate(rows[1:], start=2):
        cs = [_cell("%s%d" % (L[j], i), r[j]) for j in range(13)]
        for j, col in enumerate("NOPQRSTUV"):
            cs.append(_cell("%s%d" % (col, i), r[13 + j], forms[col].format(n=i)))
        body.append('<row r="%d" spans="1:22">%s</row>' % (i, "".join(cs)))
    new = sx[:a] + "<sheetData>" + row1 + "".join(body) + sx[b:]
    new = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:V%d"/>' % n, new, 1)

    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r"<calcPr ", '<calcPr fullCalcOnLoad="1" ', wbx, 1)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == path:
                data = new.encode("utf8")
            elif it.filename == "xl/workbook.xml":
                data = wbx.encode("utf8")
            z.writestr(it, data)
    out.seek(0)
    return out, {"rows": n - 1}
