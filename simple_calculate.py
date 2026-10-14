# In VS Code, run this once in the terminal:
#   pip install numpy pandas matplotlib openpyxl
# Then open this file and click Run Python File.
#
# pandas reads ref.xlsx.
# numpy calculates OEE, breakdown frequency, MTTR and MTBF.
# matplotlib draws the chart.
# The result workbook and the chart are saved on your Desktop.

import os
from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd

if not os.environ.get("DISPLAY"):
    matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def desktop_folder():
    home = Path.home()
    for folder in (
        home / "Desktop",
        home / "OneDrive" / "Desktop",
        home / "OneDrive - Personal" / "Desktop",
    ):
        if folder.is_dir():
            return folder
    folder = home / "Desktop"
    folder.mkdir(exist_ok=True)
    return folder


def find_workbook():
    beside_script = Path(__file__).resolve().parent / "ref.xlsx"
    if beside_script.is_file():
        return beside_script
    in_this_folder = Path("ref.xlsx")
    if in_this_folder.is_file():
        return in_this_folder.resolve()
    raise FileNotFoundError("Put ref.xlsx in the same folder as this script.")


def oee(availability, performance, quality):
    """OEE = availability x performance x quality. Inputs are percentages."""
    availability = np.asarray(availability, dtype=float)
    performance = np.asarray(performance, dtype=float)
    quality = np.asarray(quality, dtype=float)
    return availability / 100 * performance / 100 * quality / 100 * 100


table = pd.read_excel(find_workbook())
table.columns = [str(name).strip().lower() for name in table.columns]

dates = pd.to_datetime(table["date"])
operating_hours = table["operating hours"].to_numpy(dtype=float)
breakdown_hours = table["breakdown time"].to_numpy(dtype=float)
incidents = table["no.of incidents=no.of failures"].to_numpy(dtype=float)
recorded_oee = table["oee"].to_numpy(dtype=float)

blank = np.isnan(operating_hours) | np.isnan(recorded_oee)
dates = dates.to_numpy()[~blank]
operating_hours = operating_hours[~blank]
breakdown_hours = np.nan_to_num(breakdown_hours[~blank])
incidents = np.nan_to_num(incidents[~blank])
recorded_oee = recorded_oee[~blank]

breakdown_frequency = np.divide(
    incidents,
    operating_hours,
    out=np.full(incidents.shape, np.nan),
    where=operating_hours != 0,
)
mttr = np.divide(
    breakdown_hours,
    incidents,
    out=np.full(incidents.shape, np.nan),
    where=incidents != 0,
)
mtbf = np.divide(
    operating_hours,
    incidents,
    out=np.full(incidents.shape, np.nan),
    where=incidents != 0,
)

month_oee = np.mean(recorded_oee)
month_frequency = np.sum(incidents) / np.sum(operating_hours)
month_mttr = np.sum(breakdown_hours) / np.sum(incidents)
month_mtbf = np.sum(operating_hours) / np.sum(incidents)

results = pd.DataFrame(
    {
        "date": dates,
        "operating_hours": operating_hours,
        "breakdown_hours": breakdown_hours,
        "incidents": incidents,
        "oee": recorded_oee,
        "breakdown_frequency": breakdown_frequency,
        "mttr": mttr,
        "mtbf": mtbf,
    }
)
month = pd.DataFrame(
    [
        {
            "oee": month_oee,
            "breakdown_frequency": month_frequency,
            "mttr": month_mttr,
            "mtbf": month_mtbf,
            "breakdowns": np.sum(incidents),
            "operating_hours": np.sum(operating_hours),
            "breakdown_hours": np.sum(breakdown_hours),
        }
    ]
)

print("Example OEE: availability 80, performance 90, quality 95 =", round(float(oee(80, 90, 95)), 2))
print()
print(results.to_string(index=False, float_format=lambda value: f"{value:.4f}"))
print()
print(month.round(4).to_string(index=False))

desktop = desktop_folder()
workbook_out = desktop / "maintenance_results.xlsx"
chart_out = desktop / "maintenance_chart.png"

with pd.ExcelWriter(workbook_out) as writer:
    results.to_excel(writer, sheet_name="Daily", index=False)
    month.to_excel(writer, sheet_name="Month", index=False)

figure, axes = plt.subplots(3, 1, figsize=(9, 8), sharex=True)
axes[0].plot(dates, recorded_oee, marker="o", label="Daily OEE")
axes[0].axhline(82, linestyle="--", label="Year 1 target 82%")
axes[0].set_ylabel("OEE (%)")
axes[0].set_title("Daily OEE")
axes[0].legend()

axes[1].bar(dates, breakdown_frequency, width=0.8)
axes[1].set_ylabel("Breakdowns per operating hour")
axes[1].set_title("Breakdown frequency")

axes[2].plot(dates, mttr, marker="o", label="MTTR")
axes[2].plot(dates, mtbf, marker="o", label="MTBF")
axes[2].set_ylabel("Hours")
axes[2].set_xlabel("September 2026")
axes[2].set_title("MTTR and MTBF")
axes[2].legend()

start = dates.min() - np.timedelta64(12, "h")
end = dates.max() + np.timedelta64(12, "h")
for axis in axes:
    axis.set_xlim(start, end)
axes[2].xaxis.set_major_locator(mdates.DayLocator(bymonthday=range(1, 20, 2)))
axes[2].xaxis.set_major_formatter(mdates.DateFormatter("%d"))

figure.autofmt_xdate(rotation=0, ha="center")
figure.tight_layout()
figure.savefig(chart_out, dpi=120)

print()
print("Saved", workbook_out)
print("Saved", chart_out)
if matplotlib.get_backend().lower() != "agg":
    plt.show()
