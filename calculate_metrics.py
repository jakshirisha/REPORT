#!/usr/bin/env python3
"""Calculate OEE, breakdown frequency, MTTR and MTBF.

Run this file in VS Code:

1. Install Python from https://www.python.org if it is not already installed.
2. In VS Code, install the Microsoft extension named Python.
3. Use File, Open Folder, and open the folder that contains this file and ref.xlsx.
4. Open calculate_metrics.py.
5. Click Run Python File. That is the triangle at the top right of the editor.
   You can also press F5 and choose "Run maintenance metrics".

The results appear in the Terminal panel at the bottom.

OEE                    = availability x performance x quality
Breakdown frequency    = number of breakdowns / operating hours
MTTR                   = breakdown hours / number of incidents
MTBF                   = operating hours / number of failures

Percentages are on a 0–100 scale, so 82 means 82%.
"""

from __future__ import annotations

import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

EXCEL_EPOCH = datetime(1899, 12, 30)
SHEET_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def oee(availability, performance, quality):
    """Overall equipment effectiveness, in percent."""
    return (availability / 100) * (performance / 100) * (quality / 100) * 100


def breakdown_frequency(number_of_breakdowns, operating_hours):
    """Breakdowns per operating hour."""
    if operating_hours == 0:
        raise ValueError("operating hours is 0")
    return number_of_breakdowns / operating_hours


def mttr(breakdown_hours, number_of_incidents):
    """Mean time to repair, in hours."""
    if number_of_incidents == 0:
        raise ValueError("number of incidents is 0")
    return breakdown_hours / number_of_incidents


def mtbf(operating_hours, number_of_failures):
    """Mean time between failures, in hours."""
    if number_of_failures == 0:
        raise ValueError("number of failures is 0")
    return operating_hours / number_of_failures


def calculate_sheet(path):
    """Daily and monthly figures from the workbook.

    The sheet stores a recorded OEE for each day. Performance is not in the
    file, so the OEE formula is not reapplied to those rows. Call oee()
    when availability, performance and quality are all known.
    """
    daily = []
    for row in _load_rows(path):
        incidents = row["incidents"]
        hours = row["operating_hours"]
        breakdown_hours = row["breakdown_hours"]
        daily.append(
            {
                "date": row["date"],
                "oee": row["oee"],
                "breakdown_frequency": breakdown_frequency(incidents, hours) if hours else None,
                "mttr": mttr(breakdown_hours, incidents) if incidents else None,
                "mtbf": mtbf(hours, incidents) if incidents else None,
                "operating_hours": hours,
                "breakdown_hours": breakdown_hours,
                "incidents": incidents,
            }
        )

    operating = sum(item["operating_hours"] for item in daily)
    breakdown = sum(item["breakdown_hours"] for item in daily)
    incidents = sum(item["incidents"] for item in daily)
    oee_values = [item["oee"] for item in daily if item["oee"] is not None]
    month = {
        "days": len(daily),
        "oee": sum(oee_values) / len(oee_values) if oee_values else None,
        "breakdown_frequency": breakdown_frequency(incidents, operating),
        "breakdowns": incidents,
        "mttr": mttr(breakdown, incidents) if incidents else None,
        "mtbf": mtbf(operating, incidents) if incidents else None,
        "operating_hours": operating,
        "breakdown_hours": breakdown,
    }
    return daily, month


def _load_rows(path):
    table = _read_xlsx(Path(path))
    if not table:
        raise ValueError(f"{path} has no rows")
    headers = [_header(cell) for cell in table[0]]
    columns = {}
    for index, header in enumerate(headers):
        if header in {"date", "operating_hours", "breakdown_hours", "incidents", "oee"}:
            columns.setdefault(header, index)
    missing = [name for name in ("operating_hours", "breakdown_hours", "incidents") if name not in columns]
    if missing:
        raise ValueError(f"{path} is missing columns: {', '.join(missing)}")

    rows = []
    for raw in table[1:]:
        if not any(cell.strip() for cell in raw):
            continue
        rows.append(
            {
                "date": _date(raw, columns.get("date")),
                "operating_hours": _number(raw, columns.get("operating_hours")) or 0,
                "breakdown_hours": _number(raw, columns.get("breakdown_hours")) or 0,
                "incidents": _number(raw, columns.get("incidents")) or 0,
                "oee": _number(raw, columns.get("oee")),
            }
        )
    if not rows:
        raise ValueError(f"{path} has headers but no daily rows")
    return rows


def _read_xlsx(path):
    with zipfile.ZipFile(path) as workbook:
        shared = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in root.findall("m:si", SHEET_NS):
                shared.append(
                    "".join(
                        node.text or ""
                        for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    )
                )
        sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))

    rows = []
    for row in sheet.findall("m:sheetData/m:row", SHEET_NS):
        values = {}
        for cell in row.findall("m:c", SHEET_NS):
            ref = cell.get("r") or "A1"
            letters = "".join(char for char in ref if char.isalpha())
            column = 0
            for char in letters:
                column = column * 26 + (ord(char.upper()) - 64)
            kind = cell.get("t")
            value = cell.find("m:v", SHEET_NS)
            text = ""
            if kind == "s" and value is not None and value.text:
                text = shared[int(value.text)]
            elif value is not None and value.text:
                text = value.text
            values[column] = text
        if not values:
            rows.append([])
            continue
        width = max(values)
        rows.append([values.get(index, "") for index in range(1, width + 1)])
    return rows


def _header(label):
    compact = "".join(char for char in label.lower() if char.isalnum())
    return {
        "date": "date",
        "operatinghours": "operating_hours",
        "breakdowntime": "breakdown_hours",
        "noofincidentsnooffailures": "incidents",
        "noofincidents": "incidents",
        "incidents": "incidents",
        "oee": "oee",
    }.get(compact, compact)


def _number(raw, index):
    if index is None or index >= len(raw) or not raw[index].strip():
        return None
    return float(raw[index])


def _date(raw, index):
    if index is None or index >= len(raw) or not raw[index].strip():
        return None
    text = raw[index].strip()
    if text.isdigit():
        return (EXCEL_EPOCH + timedelta(days=int(text))).date().isoformat()
    return text


def _report(path, daily, month):
    lines = [
        f"Source: {path}",
        "",
        f"{'Date':<12} {'OEE %':>8} {'Breakdown frequency':>22} {'MTTR h':>8} {'MTBF h':>8}",
    ]
    for item in daily:
        lines.append(
            f"{item['date'] or '':<12} "
            f"{_fmt(item['oee'], 1):>8} "
            f"{_fmt(item['breakdown_frequency'], 4):>22} "
            f"{_fmt(item['mttr'], 2):>8} "
            f"{_fmt(item['mtbf'], 2):>8}"
        )
    lines.extend(
        [
            "",
            f"Days                         {month['days']}",
            f"OEE, average of daily OEE    {_fmt(month['oee'], 2)}%",
            f"Breakdown frequency          {_fmt(month['breakdown_frequency'], 4)} per operating hour",
            f"Breakdowns                   {month['breakdowns']:g}",
            f"MTTR                         {_fmt(month['mttr'], 2)} h",
            f"MTBF                         {_fmt(month['mtbf'], 2)} h",
            "",
            "Month formulas",
            f"  OEE = sum of daily OEE / {month['days']}",
            f"  Breakdown frequency = {month['breakdowns']:g} / {month['operating_hours']:g} operating hours",
            f"  MTTR = {month['breakdown_hours']:g} breakdown hours / {month['breakdowns']:g} incidents",
            f"  MTBF = {month['operating_hours']:g} operating hours / {month['breakdowns']:g} failures",
        ]
    )
    return "\n".join(lines)


def _fmt(value, places):
    if value is None:
        return "—"
    return f"{value:.{places}f}"


def workbook_path(argument=None):
    """Find ref.xlsx beside this script, which is where VS Code's Run button looks."""
    if argument:
        return argument
    beside_script = Path(__file__).resolve().parent / "ref.xlsx"
    if beside_script.exists():
        return str(beside_script)
    return "ref.xlsx"


def main(argv=None):
    arguments = list(sys.argv[1:] if argv is None else argv)
    path = workbook_path(arguments[0] if arguments else None)
    daily, month = calculate_sheet(path)
    print(_report(path, daily, month))
    return 0


if __name__ == "__main__":
    sys.exit(main())
