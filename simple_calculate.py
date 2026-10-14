# How to run in VS Code
# 1. Put this file and ref.xlsx in the same folder.
# 2. Click Run Python File (the triangle at the top right).
#
# This file does not load pandas. Python 3.14 on your PC freezes during
# "import pandas", which is the KeyboardInterrupt you have been seeing.
# The calculations are the same. The result files are saved beside this file.

import csv
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

EXCEL_EPOCH = datetime(1899, 12, 30)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_sheet(path):
    """Read the first sheet of an Excel file. Returns a list of rows."""
    with zipfile.ZipFile(path) as workbook:
        shared = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in root.findall("m:si", NS):
                shared.append("".join(node.text or "" for node in item.iter() if node.tag.endswith("}t")))
        sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))

    rows = []
    for row in sheet.findall("m:sheetData/m:row", NS):
        values = {}
        for cell in row.findall("m:c", NS):
            letters = "".join(ch for ch in (cell.get("r") or "A1") if ch.isalpha())
            column = 0
            for ch in letters:
                column = column * 26 + (ord(ch.upper()) - 64)
            value = cell.find("m:v", NS)
            text = ""
            if cell.get("t") == "s" and value is not None and value.text:
                text = shared[int(value.text)]
            elif value is not None and value.text:
                text = value.text
            values[column] = text
        if values:
            width = max(values)
            rows.append([values.get(i, "") for i in range(1, width + 1)])
    return rows


def number(text):
    text = text.strip()
    if text == "":
        return 0.0
    return float(text)


# Step 1. Read ref.xlsx from the same folder as this script.
excel_file = Path(__file__).resolve().parent / "ref.xlsx"
sheet = read_sheet(excel_file)
headers = [name.strip().lower() for name in sheet[0]]
column = {name: index for index, name in enumerate(headers)}

dates = []
operating_hours = []
breakdown_hours = []
incidents = []
oee = []

for raw in sheet[1:]:
    if column["oee"] >= len(raw) or raw[column["oee"]].strip() == "":
        continue
    serial = int(float(raw[column["date"]]))
    dates.append(EXCEL_EPOCH + timedelta(days=serial))
    operating_hours.append(number(raw[column["operating hours"]]))
    breakdown_hours.append(number(raw[column["breakdown time"]]))
    incidents.append(number(raw[column["no.of incidents=no.of failures"]]))
    oee.append(number(raw[column["oee"]]))


# Step 2. Calculate one value for every day.
# Breakdown frequency = incidents / operating hours
# MTTR = breakdown hours / incidents
# MTBF = operating hours / incidents
# A day with 0 incidents has no MTTR and no MTBF.

breakdown_frequency = []
mttr = []
mtbf = []

for i in range(len(dates)):
    breakdown_frequency.append(incidents[i] / operating_hours[i])
    if incidents[i] == 0:
        mttr.append(None)
        mtbf.append(None)
    else:
        mttr.append(breakdown_hours[i] / incidents[i])
        mtbf.append(operating_hours[i] / incidents[i])


# Step 3. Calculate the whole month from the totals.
month_oee = sum(oee) / len(oee)
month_breakdown_frequency = sum(incidents) / sum(operating_hours)
month_mttr = sum(breakdown_hours) / sum(incidents)
month_mtbf = sum(operating_hours) / sum(incidents)

print("Date        OEE    Frequency    MTTR    MTBF")
for i in range(len(dates)):
    mttr_text = "-" if mttr[i] is None else f"{mttr[i]:.2f}"
    mtbf_text = "-" if mtbf[i] is None else f"{mtbf[i]:.2f}"
    print(
        dates[i].strftime("%Y-%m-%d"),
        f"{oee[i]:.1f}",
        f"{breakdown_frequency[i]:.4f}",
        mttr_text,
        mtbf_text,
    )

print()
print("Month OEE                ", round(month_oee, 2))
print("Month breakdown frequency", round(month_breakdown_frequency, 4))
print("Month MTTR               ", round(month_mttr, 2))
print("Month MTBF               ", round(month_mtbf, 2))


# Step 4. Save the numbers next to this script (your Desktop).
folder = Path(__file__).resolve().parent
results_file = folder / "maintenance_results.csv"
with results_file.open("w", newline="") as handle:
    writer = csv.writer(handle)
    writer.writerow(["Date", "operating hours", "breakdown time", "incidents", "OEE", "breakdown frequency", "MTTR", "MTBF"])
    for i in range(len(dates)):
        writer.writerow([
            dates[i].strftime("%Y-%m-%d"),
            operating_hours[i],
            breakdown_hours[i],
            incidents[i],
            oee[i],
            round(breakdown_frequency[i], 4),
            "" if mttr[i] is None else round(mttr[i], 2),
            "" if mtbf[i] is None else round(mtbf[i], 2),
        ])
    writer.writerow([])
    writer.writerow(["Month OEE", round(month_oee, 2)])
    writer.writerow(["Month breakdown frequency", round(month_breakdown_frequency, 4)])
    writer.writerow(["Month MTTR", round(month_mttr, 2)])
    writer.writerow(["Month MTBF", round(month_mtbf, 2)])

print()
print("Saved", results_file)
