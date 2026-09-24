#!/usr/bin/env python3
"""Funnel a breakdown JSON (exported from the VFX Script Breakdown tool) into a
bid workbook (.xlsx).

Usage:
    python3 pipeline/build_xlsx.py MY_FILM_breakdown.json [output.xlsx]

Copies the bundled workbook template (template.xlsx), fills the Shot Breakdown
sheet with your rows, and writes live Excel formulas so totals, averages, and
the Sequence Breakdown stay linked. The project / sequence codes come from the
JSON (set them in the tool's Export dialog).

Requires: pip install openpyxl
"""
import base64
import copy as _copy
import io
import json
import shutil
import sys
import uuid
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

TEMPLATE = Path(__file__).parent / "template.xlsx"
DATA_START = 6  # first data row in Shot Breakdown


def main():
    src = Path(sys.argv[1])
    data = json.loads(src.read_text())
    proj = (data.get("proj") or "PRJ").strip() or "PRJ"
    seq = (data.get("seq") or "PRJ").strip() or "PRJ"
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(
        proj + "_VFX_Breakdown.xlsx")

    include_prices = data.get("includePrices", True)

    shutil.copy(TEMPLATE, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Shot Breakdown"]

    # one estimate column ("Days (est)") ahead of cost. Insert first so every
    # later write uses the shifted layout:
    #   N=shots  O=Days  P=cost/shot  Q=total  R=image
    hdr_style = _copy.copy(ws.cell(5, 15)._style)
    ws.insert_cols(15, 1)
    ws.cell(5, 15, "Days (est)")._style = _copy.copy(hdr_style)
    ws.column_dimensions["O"].width = 9

    # capture the template's own data-row formatting so generated rows match
    proto = [_copy.copy(ws.cell(DATA_START, c)._style) for c in range(1, 21)]

    ws.delete_rows(DATA_START, max(1, ws.max_row - DATA_START + 1))

    widest = [0]

    def put(row, col, value=None):
        c = ws.cell(row, col, value)
        c._style = _copy.copy(proto[col - 1])
        return c

    ws["A1"] = data["title"]

    rows = data["rows"]
    r = DATA_START
    for it in rows:
        is_vfx = it["status"] == "vfx"
        scope = ("\n".join(it["elements"]) if is_vfx and it["elements"]
                 else "VFX TBD" if it["status"] == "tbd"
                 else "SCENE NOTES" if it["status"] == "note"
                 else "NO VFX")
        notes = ". ".join(x for x in [
            "Method: " + " / ".join(it["methods"]) if it.get("methods") else "",
            it.get("notes", ""),
        ] if x)

        for col in range(1, 21):
            put(r, col)
        put(r, 1, str(uuid.uuid4()))
        put(r, 2, it["item"])
        put(r, 3, it["page"])
        put(r, 4, str(it["scene"]))
        put(r, 5, seq)
        put(r, 6, it.get("ie", ""))
        put(r, 7, it.get("location", ""))
        put(r, 8, it.get("tod", ""))
        put(r, 9, it.get("char", ""))
        put(r, 10, it["text"])
        put(r, 11, notes)
        put(r, 12, scope).alignment = Alignment(wrap_text=True, vertical="center")
        if is_vfx:
            put(r, 13, it.get("lod", ""))
            put(r, 14, it.get("shots", 0))
            put(r, 15, it.get("days") or "")
            if include_prices:
                put(r, 16, it.get("cost", 0))
                put(r, 17, f"=N{r}*P{r}")
        elif include_prices:
            put(r, 17, 0)
        if it.get("img", "").startswith("data:image"):
            try:
                from openpyxl.drawing.image import Image as XLImage
                raw = base64.b64decode(it["img"].split(",", 1)[1])
                pic = XLImage(io.BytesIO(raw))
                px_h = 120
                scale = px_h / pic.height
                pic.height, pic.width = px_h, int(pic.width * scale)
                pic.anchor = f"R{r}"
                ws.add_image(pic)
                ws.row_dimensions[r].height = px_h * 0.75 + 6
                widest[0] = max(widest[0], pic.width)
            except Exception:
                ws.cell(r, 16, "[storyboard in tool]")
        r += 1

    last = r - 1
    # allowance row
    allow_shots = int(data.get("allowShots") or 0)
    allow_rate = int(data.get("allowRate") or 0)
    if allow_shots:
        for col in range(1, 21):
            put(r, col)
        put(r, 1, str(uuid.uuid4()))
        put(r, 5, "ALLOWANCE")
        put(r, 10, "VFX Allowance")
        put(r, 11, "Unassigned allowance shots")
        put(r, 12, "VFX Allowance")
        put(r, 14, allow_shots)
        if include_prices:
            put(r, 16, allow_rate)
            put(r, 17, f"=N{r}*P{r}")

    # size the storyboard column to the widest board
    if widest[0]:
        ws.column_dimensions["R"].width = max(
            ws.column_dimensions["R"].width or 0, widest[0] / 7.0 + 2)

    ws.auto_filter.ref = f"A5:R{r if allow_shots else last}"

    # live summary block (top-right)
    ws["N2"] = f"=SUM(N{DATA_START}:N{last})"
    ws["P2"] = "=Q2/N2"
    ws["Q2"] = f"=SUM(Q{DATA_START}:Q{last})"
    ws["S2"] = None
    ws["T2"] = None
    if allow_shots:
        ws["N3"], ws["P3"], ws["Q3"] = allow_shots, allow_rate, "=N3*P3"
    else:
        ws["N3"], ws["P3"], ws["Q3"] = 0, 0, 0
    ws["N4"] = "=N2+N3"
    ws["Q4"] = "=Q2+Q3"

    if not include_prices:
        for ref in ("P2", "Q2", "P3", "Q3", "Q4", "S2", "T2"):
            ws[ref] = None

    # sequence breakdown sheet
    sq = wb["Sequence Breakdown"]
    sq["A1"] = data["title"]
    pages = max((it["page"] for it in rows), default=0)
    sq["A4"], sq["B4"], sq["C4"] = 1 + (1 if allow_shots else 0), pages, "=C6+C7"
    sq["A6"], sq["B6"] = seq, pages
    sq["C6"] = "='Shot Breakdown'!N2"
    sq["D6"] = "='Shot Breakdown'!Q2"
    sq["A7"] = "ALLOWANCE" if allow_shots else ""
    sq["C7"] = allow_shots or ""
    sq["D7"] = "='Shot Breakdown'!Q3" if allow_shots else ""

    # show summary header + counts
    if "Show Summary" in wb.sheetnames:
        sm = wb["Show Summary"]
        sm["A1"] = data["title"]
        sm["A2"] = ""
        sm["D8"] = sum(1 for it in rows if it["status"] == "vfx")
        sm["D9"] = allow_shots or 0

    wb.save(out)
    n_vfx = sum(1 for it in rows if it["status"] == "vfx")
    shots = sum(it.get("shots", 0) for it in rows if it["status"] == "vfx")
    days = sum((it.get("shots", 0) or 0) * (it.get("days", 0) or 0)
               for it in rows if it["status"] == "vfx")
    cost = sum(it.get("shots", 0) * it.get("cost", 0) for it in rows if it["status"] == "vfx")
    n_img = len(ws._images)
    print(f"{out.name}: {len(rows)} rows, {n_vfx} VFX items, {shots} shots, "
          f"{days:g} est. days, ${cost:,} (+ ${allow_shots*allow_rate:,} allowance), "
          f"{n_img} storyboard{'' if n_img == 1 else 's'} embedded"
          + (" -> Shot Breakdown tab, column R" if n_img else ""))


if __name__ == "__main__":
    main()
