# OEE = availability x performance x quality
# Breakdown frequency = breakdowns / operating hours
# MTTR = breakdown hours / incidents
# MTBF = operating hours / failures


def oee(availability, performance, quality):
    return (availability / 100) * (performance / 100) * (quality / 100) * 100


def breakdown_frequency(breakdowns, operating_hours):
    return breakdowns / operating_hours


def mttr(breakdown_hours, incidents):
    return breakdown_hours / incidents


def mtbf(operating_hours, failures):
    return operating_hours / failures


# date, operating hours, breakdown hours, incidents, recorded OEE
rows = [
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
]

print("Example OEE: availability 80, performance 90, quality 95 =", round(oee(80, 90, 95), 2))
print()
print("Date          OEE    Frequency    MTTR    MTBF")

total_hours = 0
total_breakdown = 0
total_incidents = 0
total_oee = 0

for date, hours, breakdown, incidents, recorded_oee in rows:
    total_hours += hours
    total_breakdown += breakdown
    total_incidents += incidents
    total_oee += recorded_oee

    frequency = breakdown_frequency(incidents, hours)
    repair_time = mttr(breakdown, incidents) if incidents else None
    between_failures = mtbf(hours, incidents) if incidents else None

    repair_text = "-" if repair_time is None else round(repair_time, 2)
    between_text = "-" if between_failures is None else round(between_failures, 2)
    print(date, round(recorded_oee, 1), round(frequency, 4), repair_text, between_text)

print()
print("Month OEE               ", round(total_oee / len(rows), 2))
print("Month breakdown frequency", round(breakdown_frequency(total_incidents, total_hours), 4))
print("Month MTTR              ", round(mttr(total_breakdown, total_incidents), 2))
print("Month MTBF              ", round(mtbf(total_hours, total_incidents), 2))
