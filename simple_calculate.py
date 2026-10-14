# In VS Code, run this once in the terminal:
#   pip install numpy pandas matplotlib
# Then open this file and click Run Python File.

import os
from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd

if not os.environ.get("DISPLAY"):
    matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def oee(availability, performance, quality):
    """OEE = availability x performance x quality. Inputs are percentages."""
    availability = np.asarray(availability, dtype=float)
    performance = np.asarray(performance, dtype=float)
    quality = np.asarray(quality, dtype=float)
    return availability / 100 * performance / 100 * quality / 100 * 100


# date, operating hours, breakdown hours, incidents, recorded OEE
data = pd.DataFrame(
    [
        ("2026-09-01", 25, 5, 2, 85.9),
        ("2026-09-02", 23, 6, 3, 89),
        ("2026-09-03", 26, 4, 2, 87),
        ("2026-09-04", 23, 2, 1, 92),
        ("2026-09-05", 24, 1, 1, 82),
        ("2026-09-06", 27, 4, 3, 74),
        ("2026-09-07", 22, 2, 1, 83),
        ("2026-09-08", 28, 4, 1, 93),
        ("2026-09-09", 29, 3, 2, 81),
        ("2026-09-10", 30, 0, 0, 90),
        ("2026-09-11", 40, 7, 4, 94),
        ("2026-09-12", 50, 10, 5, 88),
        ("2026-09-13", 45, 6, 4, 89),
        ("2026-09-14", 50, 6, 4, 89),
        ("2026-09-15", 67, 5, 4, 85),
        ("2026-09-16", 35, 0, 0, 94),
        ("2026-09-17", 62, 0, 0, 96),
        ("2026-09-18", 45, 0, 0, 87),
        ("2026-09-19", 35, 2, 0, 98),
    ],
    columns=["date", "operating_hours", "breakdown_hours", "incidents", "oee"],
)
data["date"] = pd.to_datetime(data["date"])

data["breakdown_frequency"] = np.divide(data["incidents"], data["operating_hours"])
data["mttr"] = np.divide(data["breakdown_hours"], data["incidents"].replace(0, np.nan))
data["mtbf"] = np.divide(data["operating_hours"], data["incidents"].replace(0, np.nan))

total_hours = data["operating_hours"].sum()
total_breakdown = data["breakdown_hours"].sum()
total_incidents = data["incidents"].sum()
month = pd.Series(
    {
        "oee": data["oee"].mean(),
        "breakdown_frequency": total_incidents / total_hours,
        "mttr": total_breakdown / total_incidents,
        "mtbf": total_hours / total_incidents,
    }
)

print("Example OEE: availability 80, performance 90, quality 95 =", round(float(oee(80, 90, 95)), 2))
print()
print(data.to_string(index=False, float_format=lambda value: f"{value:.4f}"))
print()
print(month.round(4).to_string())

figure, axes = plt.subplots(3, 1, figsize=(9, 8), sharex=True)

axes[0].plot(data["date"], data["oee"], marker="o", label="Daily OEE")
axes[0].axhline(82, linestyle="--", label="Year 1 target 82%")
axes[0].set_ylabel("OEE (%)")
axes[0].set_title("Daily OEE, 1–19 September 2026")
axes[0].legend()

axes[1].bar(data["date"], data["breakdown_frequency"], width=0.8)
axes[1].set_ylabel("Breakdowns per operating hour")
axes[1].set_title("Breakdown frequency = incidents / operating hours")

axes[2].plot(data["date"], data["mttr"], marker="o", label="MTTR")
axes[2].plot(data["date"], data["mtbf"], marker="o", label="MTBF")
axes[2].set_ylabel("Hours")
axes[2].set_xlabel("September 2026")
axes[2].set_title("MTTR and MTBF")
axes[2].legend()
start = data["date"].min() - pd.Timedelta(hours=12)
end = data["date"].max() + pd.Timedelta(hours=12)
for axis in axes:
    axis.set_xlim(start, end)
axes[2].xaxis.set_major_locator(mdates.DayLocator(bymonthday=range(1, 20, 2)))
axes[2].xaxis.set_major_formatter(mdates.DateFormatter("%d"))

figure.autofmt_xdate(rotation=0, ha="center")
figure.tight_layout()
chart = Path(__file__).resolve().parent / "maintenance_chart.png"
figure.savefig(chart, dpi=120)
print()
print("Chart saved to", chart.name)
if matplotlib.get_backend().lower() != "agg":
    plt.show()
