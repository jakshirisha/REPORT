# How to run in VS Code
# 1. In the terminal, type:  pip install numpy pandas matplotlib openpyxl
# 2. Keep ref.xlsx in the same folder as this file.
# 3. Click Run Python File (the triangle at the top right).
#
# pandas reads the Excel file.
# numpy holds the numbers and does the sums and divisions.
# matplotlib draws the chart.
# The two result files are saved on your Desktop.

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# Step 1. Read the Excel file.
excel_file = Path(__file__).resolve().parent / "ref.xlsx"
table = pd.read_excel(excel_file)

# The sheet has a few empty rows at the bottom. Drop them.
table = table.dropna(subset=["OEE"])


# Step 2. Copy the columns into numpy arrays.
dates = table["Date"].to_numpy()
operating_hours = table["operating hours"].to_numpy(dtype=float)
breakdown_hours = table["breakdown time"].to_numpy(dtype=float)
incidents = table["No.of incidents=No.of Failures"].to_numpy(dtype=float)
oee = table["OEE"].to_numpy(dtype=float)


# Step 3. Calculate one value for every day.
# Breakdown frequency = incidents / operating hours
# MTTR = breakdown hours / incidents
# MTBF = operating hours / incidents
# A day with 0 incidents has no MTTR and no MTBF.

day_count = len(dates)
breakdown_frequency = np.empty(day_count)
mttr = np.empty(day_count)
mtbf = np.empty(day_count)

for i in range(day_count):
    breakdown_frequency[i] = incidents[i] / operating_hours[i]

    if incidents[i] == 0:
        mttr[i] = np.nan
        mtbf[i] = np.nan
    else:
        mttr[i] = breakdown_hours[i] / incidents[i]
        mtbf[i] = operating_hours[i] / incidents[i]


# Step 4. Calculate the whole month from the totals.
month_oee = np.mean(oee)
month_breakdown_frequency = np.sum(incidents) / np.sum(operating_hours)
month_mttr = np.sum(breakdown_hours) / np.sum(incidents)
month_mtbf = np.sum(operating_hours) / np.sum(incidents)

print("OEE formula check: 80 x 90 x 95 =", round((80 / 100) * (90 / 100) * (95 / 100) * 100, 2))
print()
print("Date        OEE    Frequency    MTTR    MTBF")
for i in range(day_count):
    date_text = pd.Timestamp(dates[i]).strftime("%Y-%m-%d")
    frequency_text = f"{breakdown_frequency[i]:.4f}"
    mttr_text = "-" if np.isnan(mttr[i]) else f"{mttr[i]:.2f}"
    mtbf_text = "-" if np.isnan(mtbf[i]) else f"{mtbf[i]:.2f}"
    print(date_text, f"{oee[i]:.1f}", frequency_text, mttr_text, mtbf_text)

print()
print("Month OEE                ", round(month_oee, 2))
print("Month breakdown frequency", round(month_breakdown_frequency, 4))
print("Month MTTR               ", round(month_mttr, 2))
print("Month MTBF               ", round(month_mtbf, 2))


# Step 5. Save the numbers on the Desktop.
desktop = Path.home() / "Desktop"
if not desktop.is_dir():
    desktop = Path.home() / "OneDrive" / "Desktop"
desktop.mkdir(parents=True, exist_ok=True)

daily_result = pd.DataFrame({
    "Date": dates,
    "operating hours": operating_hours,
    "breakdown time": breakdown_hours,
    "incidents": incidents,
    "OEE": oee,
    "breakdown frequency": breakdown_frequency,
    "MTTR": mttr,
    "MTBF": mtbf,
})
month_result = pd.DataFrame({
    "OEE": [month_oee],
    "breakdown frequency": [month_breakdown_frequency],
    "MTTR": [month_mttr],
    "MTBF": [month_mtbf],
})

results_file = desktop / "maintenance_results.xlsx"
with pd.ExcelWriter(results_file) as writer:
    daily_result.to_excel(writer, sheet_name="Daily", index=False)
    month_result.to_excel(writer, sheet_name="Month", index=False)


# Step 6. Draw the chart and save it on the Desktop.
plt.figure(figsize=(9, 8))

plt.subplot(3, 1, 1)
plt.plot(dates, oee, marker="o", label="OEE")
plt.axhline(82, linestyle="--", label="Target 82%")
plt.ylabel("OEE (%)")
plt.title("OEE")
plt.legend()

plt.subplot(3, 1, 2)
plt.bar(dates, breakdown_frequency)
plt.ylabel("Breakdowns per hour")
plt.title("Breakdown frequency")

plt.subplot(3, 1, 3)
plt.plot(dates, mttr, marker="o", label="MTTR")
plt.plot(dates, mtbf, marker="o", label="MTBF")
plt.ylabel("Hours")
plt.xlabel("Date")
plt.title("MTTR and MTBF")
plt.legend()

for picture in plt.gcf().axes:
    for label in picture.get_xticklabels():
        label.set_rotation(45)

plt.tight_layout()
chart_file = desktop / "maintenance_chart.png"
plt.savefig(chart_file)
print()
print("Saved", results_file)
print("Saved", chart_file)
plt.show()
