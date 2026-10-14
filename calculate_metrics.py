#!/usr/bin/env python3
"""Calculate OEE, breakdown frequency, MTTR and MTBF.

OEE                    = availability x performance x quality
Breakdown frequency    = number of breakdowns / operating hours
MTTR                   = breakdown hours / number of incidents
MTBF                   = operating hours / number of failures

Inputs are ordinary numbers. Percentages are on a 0–100 scale, so 82 means 82%.
"""

from __future__ import annotations

import sys

from calculate_kpis import load_daily_rows


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


def _num(value):
    return None if value is None else float(value)


def calculate_sheet(path):
    """Daily and monthly figures from the workbook.

    The sheet stores a recorded OEE for each day. Performance is not in the
    file, so the OEE formula is not reapplied to those rows. Call oee()
    when availability, performance and quality are all known.
    """
    daily = []
    for row in load_daily_rows(path):
        incidents = _num(row.incidents)
        hours = _num(row.operating_hours)
        breakdown_hours = _num(row.breakdown_hours)
        daily.append(
            {
                "date": None if row.date is None else row.date.date().isoformat(),
                "oee": _num(row.oee),
                "breakdown_frequency": (
                    breakdown_frequency(incidents, hours) if hours else None
                ),
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


def main(argv=None):
    arguments = list(sys.argv[1:] if argv is None else argv)
    path = arguments[0] if arguments else "ref.xlsx"
    daily, month = calculate_sheet(path)
    print(_report(path, daily, month))
    return 0


if __name__ == "__main__":
    sys.exit(main())
