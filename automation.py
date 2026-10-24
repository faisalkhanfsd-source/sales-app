"""Q-Automation workbook: rebuild the 'Automation' sheet from the source sheets.

Automation!A:M  = FILTER of 'Customer Wise'!A:M            (kept as the dynamic-array formula)
Automation!N:V  = per-row formulas (XLOOKUP / SUMIFS ...)  -> extended / trimmed to the row count
Everything else in the workbook is copied byte-for-byte; only sheet XML of 'Automation' changes.
Cached values are computed here so the file also reads correctly before Excel recalculates.
"""
import io, re, zipfile, datetime
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.utils.datetime import to_excel

SHEETS = {"cw": "Customer Wise", "sol": "Sales Order List", "pvt": "Pivet Customer Dispatch",
          "fin": "Finalized", "auto": "Automation"}
NA = "#N/A"


def _rows(wb, name, cols, first=1):
    ws = wb[name]
    out = []
    for row in ws.iter_rows(min_row=first, values_only=True):
        out.append([row[c] if c < len(row) else None for c in cols])
    return out


def _xkey(v):                      # XLOOKUP exact match: text case-insensitive, numbers as numbers
    if isinstance(v, str):
        return ("s", v.lower())
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        v = to_excel(v)
    return ("n", float(v))


def _ckey(v):                      # SUMIFS / COUNTIF criteria quirk: "00209" == 209
    if isinstance(v, str):
        t = v.strip()
        try:
            return ("n", float(t)) if re.fullmatch(r"[+-]?\d+(\.\d+)?", t) else ("s", v.lower())
        except ValueError:
            return ("s", v.lower())
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        v = to_excel(v)
    return ("n", float(v))


def _first_map(keys, vals):
    m = {}
    for k, v in zip(keys, vals):
        k = _xkey(k)
        if k is not None and k not in m:
            m[k] = v
    return m


def compute(wb):
    cw = _rows(wb, SHEETS["cw"], range(13))
    cw = [r for r in cw if r[0] not in (None, "")]            # FILTER(...A:A<>"")
    sol = _rows(wb, SHEETS["sol"], (0, 11, 12))
    pvt = _rows(wb, SHEETS["pvt"], (2, 5, 11))
    fin = _rows(wb, SHEETS["fin"], (3, 5, 6, 8))
    status = _first_map([r[0] for r in sol], [r[1] for r in sol])
    deliv = _first_map([r[0] for r in sol], [r[2] for r in sol])
    freg = _first_map([r[0] for r in fin], [r[1] for r in fin])
    fzone = _first_map([r[0] for r in fin], [r[2] for r in fin])
    fgm = _first_map([r[0] for r in fin], [r[3] for r in fin])
    disp = {}
    for a, d, q in pvt:
        if isinstance(q, (int, float)):
            disp[(_ckey(a), _ckey(d))] = disp.get((_ckey(a), _ckey(d)), 0) + q
    own = {}
    for r in cw[1:]:
        k = (_ckey(r[0]), _ckey(r[3]))
        own[k] = own.get(k, 0) + (r[6] if isinstance(r[6], (int, float)) else 0)
    out = []
    for r in cw[1:]:
        g = r[6] if isinstance(r[6], (int, float)) else 0
        p = disp.get((_ckey(r[0]), _ckey(r[3])), 0)
        rr = own.get((_ckey(r[0]), _ckey(r[3])), 0)

        def look(m, key):
            v = m.get(_xkey(key), NA)
            return 0 if v is None else v
        out.append([look(status, r[0]), look(deliv, r[0]), p, g - p, rr, g - rr,
                    look(freg, r[2]), look(fzone, r[2]), look(fgm, r[2])])
    return cw, out


def _num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return repr(float(v))


def _cell(ref, v, formula=None, extra=""):
    f = f"<f>{escape(formula)}</f>" if formula else ""
    if v is None:
        return f'<c r="{ref}"{extra}>{f}</c>' if f else ""
    if isinstance(v, datetime.datetime):
        v = to_excel(v)
    if isinstance(v, str):
        if v == NA and formula:
            return f'<c r="{ref}" t="e"{extra}>{f}<v>#N/A</v></c>'
        return f'<c r="{ref}" t="str"{extra}>{f}<v>{escape(v)}</v></c>'
    return f'<c r="{ref}"{extra}>{f}<v>{_num(v)}</v></c>'


def _sheet_files(z):
    import xml.etree.ElementTree as ET
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    tgt = {}
    for rel in rels:
        t = rel.get("Target", "")
        t = t.lstrip("/")
        tgt[rel.get("Id")] = t if t.startswith("xl/") else "xl/" + t
    rid = "{%s}id" % ns["r"]
    return {sh.get("name"): tgt[sh.get(rid)] for sh in wbx.find("m:sheets", ns)}


FORMULAS = [
    "_xlfn.XLOOKUP($A{r},'Sales Order List'!$A:$A,'Sales Order List'!L:L)",
    "_xlfn.XLOOKUP($A{r},'Sales Order List'!$A:$A,'Sales Order List'!M:M)",
    "SUMIFS('Pivet Customer Dispatch'!$L:$L,'Pivet Customer Dispatch'!$C:$C,$A{r},'Pivet Customer Dispatch'!$F:$F,D{r})",
    "G{r}-P{r}",
    "SUMIFS(G:G,A:A,A{r},D:D,D{r})",
    "G{r}-R{r}",
    "_xlfn.XLOOKUP($C{r},Finalized!$D:$D,Finalized!F:F)",
    "_xlfn.XLOOKUP($C{r},Finalized!$D:$D,Finalized!G:G)",
    "_xlfn.XLOOKUP($C{r},Finalized!$D:$D,Finalized!I:I)",
]
LET = "ABCDEFGHIJKLMNOPQRSTUV"


def rebuild(data: bytes):
    zin = zipfile.ZipFile(io.BytesIO(data))
    files = _sheet_files(zin)
    for k in SHEETS.values():
        if k not in files:
            raise ValueError(f"فائل میں '{k}' نام کی شیٹ نہیں ملی")
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    cw, calc = compute(wb)
    n = len(cw)                                          # header + data rows
    path = files[SHEETS["auto"]]
    xml = zin.read(path).decode("utf8")
    r1 = re.search(r'<row r="1".*?</row>', xml, re.S).group(0)
    head_tail = "".join(re.findall(r'<c r="[N-V]1".*?</c>', r1, re.S))
    rows = []
    hdr = [_cell("A1", cw[0][0], extra=' cm="1"').replace(
        "<v>", f'<f t="array" ref="A1:M{n}">' + escape("_xlfn._xlws.FILTER('Customer Wise'!A:M,'Customer Wise'!A:A<>\"\")") + "</f><v>", 1)]
    hdr += [_cell(f"{LET[c]}1", cw[0][c]) for c in range(1, 13)]
    rows.append(f'<row r="1" spans="1:22">{"".join(hdr)}{head_tail}</row>')
    for i in range(1, n):
        r = i + 1
        cells = [_cell(f"{LET[c]}{r}", cw[i][c]) for c in range(13)]
        cells += [_cell(f"{LET[13 + j]}{r}", calc[i - 1][j], FORMULAS[j].format(r=r)) for j in range(9)]
        rows.append(f'<row r="{r}" spans="1:22">{"".join(cells)}</row>')
    xml = re.sub(r"<sheetData>.*</sheetData>", lambda m: "<sheetData>" + "".join(rows) + "</sheetData>", xml, flags=re.S)
    xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:V{n}"/>', xml)
    wbxml = zin.read("xl/workbook.xml").decode("utf8")
    if "fullCalcOnLoad" not in wbxml:
        wbxml = wbxml.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            if item.filename == path:
                zo.writestr(item, xml)
            elif item.filename == "xl/workbook.xml":
                zo.writestr(item, wbxml)
            else:
                zo.writestr(item, zin.read(item.filename))
    return out.getvalue(), {"rows": n - 1}
