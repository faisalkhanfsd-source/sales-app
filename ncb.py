"""SALES_VS_TARGET workbook -> PNG of the SUMMARY sheet (one region block)."""
import io, os, base64, json, urllib.request
import openpyxl
from PIL import Image, ImageDraw, ImageFont

FONTS = ["/usr/share/fonts/truetype/dejavu/DejaVuSans{}.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans{}.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSans-{}.ttf"]


def _font(size, bold=False):
    for p in FONTS:
        for f in ([p.format("-Bold" if bold else "")] if "DejaVu" in p else
                  [p.format("Bold" if bold else "Regular")]):
            if os.path.exists(f):
                return ImageFont.truetype(f, size)
    try:
        return ImageFont.load_default(size)
    except TypeError:
        return ImageFont.load_default()


def read_summary(file_like):
    wb = openpyxl.load_workbook(file_like, data_only=True)
    ws = next((s for s in wb.worksheets if s.title.strip().upper() == "SUMMARY"), None)
    if ws is None:
        raise ValueError("فائل میں 'SUMMARY' نام کی شیٹ نہیں ملی")
    hdr = next((r for r in range(1, 15) if str(ws.cell(r, 1).value or "").strip().upper() == "REGION"), None)
    if hdr is None:
        raise ValueError("SUMMARY شیٹ میں REGION والی ہیڈر لائن نہیں ملی")
    date_txt = next((str(ws.cell(r, 1).value) for r in range(1, hdr)
                     if str(ws.cell(r, 1).value or "").startswith("From")), "")
    title = next((str(ws.cell(r, 1).value) for r in range(1, hdr)
                  if ws.cell(r, 1).value and not str(ws.cell(r, 1).value).startswith("From")),
                 "MTD SALE Vs Target")
    heads = [ws.cell(hdr, c).value for c in (1, 2, 4, 5, 6, 7, 8, 9)]
    regions = {}
    for r in range(hdr + 1, ws.max_row + 1):
        reg = ws.cell(r, 1).value
        if not reg:
            continue
        reg = str(reg).strip().upper()
        row = [ws.cell(r, c).value for c in (1, 2, 4, 5, 6, 7, 8, 9)]
        row[0] = str(row[0]).strip()
        regions.setdefault(reg, []).append(row)
    return {"date": date_txt, "title": title, "heads": heads, "regions": regions}


def _num(v):
    return "" if v in (None, "") else f"{float(v):,.0f}" if abs(float(v)) >= .5 else "-"


def _pct(v):
    return "" if v in (None, "") else f"{float(v) * 100:.0f}%"


def render(data, region="SBL TOTAL"):
    rows = data["regions"].get(region.upper())
    if not rows:
        raise ValueError("ریجن نہیں ملا: " + region)
    widths = [130, 250, 105, 105, 75, 105, 85, 75]
    H, pad = 34, 8
    W = sum(widths)
    n = len(rows)
    img = Image.new("RGB", (W + 2, H * (n + 3) + 2), "white")
    d = ImageDraw.Draw(img)
    f, fb = _font(15), _font(15, True)

    def cell(x, y, w, text, fill=None, bold=False, color="black", align="left"):
        d.rectangle([x, y, x + w, y + H], fill=fill or "white", outline="#9aa3b2")
        ft = fb if bold else f
        tw = d.textlength(text, font=ft)
        tx = x + (w - tw) / 2 if align == "center" else x + w - tw - pad if align == "right" else x + pad
        d.text((tx, y + H / 2), text, font=ft, fill=color, anchor="lm")

    # row 0: date (green) ; row 1: title (orange) + Target Pack Mix (yellow)
    cell(0, 0, W, data["date"], "#92D050", True, align="center")
    left = sum(widths[:6])
    cell(0, H, left, data["title"], "#FFC000", True, align="center")
    cell(left, H, W - left, "Target Pack Mix", "#FFFF00", True, align="center")
    x = 0
    for w, h in zip(widths, data["heads"]):
        cell(x, 2 * H, w, str(h), "#F2F2F2", True, align="center" if x == widths[0] else "left")
        x += w
    for i, r in enumerate(rows):
        y = (3 + i) * H
        total = not r[1]
        fill = "#002060" if total else None
        col = "white" if total else "black"
        vals = [r[0], r[1] or "", _num(r[2]), _num(r[3]), _pct(r[4]), _num(r[5]), _pct(r[6]), _pct(r[7])]
        x = 0
        for j, (w, v) in enumerate(zip(widths, vals)):
            cf = fill or ("#AAD1F4" if j == 1 else None)
            cell(x, y, w, v, cf, total, col, "left" if j < 2 else "right")
            x += w
    out = io.BytesIO()
    img.save(out, "PNG")
    return out.getvalue()


def send_whatsapp(png):
    """Optional: WhatsApp Cloud API. Needs WA_TOKEN, WA_PHONE_ID, WA_TO env vars."""
    tok, pid, to = (os.environ.get(k) for k in ("WA_TOKEN", "WA_PHONE_ID", "WA_TO"))
    if not (tok and pid and to):
        return None
    b = "----wa" + base64.b16encode(os.urandom(6)).decode()
    body = (f"--{b}\r\nContent-Disposition: form-data; name=\"messaging_product\"\r\n\r\nwhatsapp\r\n"
            f"--{b}\r\nContent-Disposition: form-data; name=\"type\"\r\n\r\nimage/png\r\n"
            f"--{b}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"report.png\"\r\n"
            f"Content-Type: image/png\r\n\r\n").encode() + png + f"\r\n--{b}--\r\n".encode()
    req = urllib.request.Request(f"https://graph.facebook.com/v20.0/{pid}/media", body,
                                 {"Authorization": "Bearer " + tok,
                                  "Content-Type": "multipart/form-data; boundary=" + b})
    mid = json.load(urllib.request.urlopen(req, timeout=60))["id"]
    msg = json.dumps({"messaging_product": "whatsapp", "to": to, "type": "image", "image": {"id": mid}}).encode()
    req = urllib.request.Request(f"https://graph.facebook.com/v20.0/{pid}/messages", msg,
                                 {"Authorization": "Bearer " + tok, "Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=60).read()
    return True
