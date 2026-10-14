# Breakdown frequency = number of breakdowns / operating hours
# Click Run Python File in VS Code.

# operating hours, number of breakdowns
rows = [
    (25, 2),
    (23, 3),
    (26, 2),
    (23, 1),
    (24, 1),
    (27, 3),
    (22, 1),
    (28, 1),
    (29, 2),
    (30, 0),
    (40, 4),
    (50, 5),
    (45, 4),
    (50, 4),
    (67, 4),
    (35, 0),
    (62, 0),
    (45, 0),
    (35, 0),
]


def breakdown_frequency(breakdowns, operating_hours):
    return breakdowns / operating_hours


print("Day   Operating hours   Breakdowns   Frequency")

total_hours = 0
total_breakdowns = 0

for day, (hours, breakdowns) in enumerate(rows, start=1):
    frequency = breakdown_frequency(breakdowns, hours)
    total_hours += hours
    total_breakdowns += breakdowns
    print(day, hours, breakdowns, round(frequency, 4))

month_frequency = breakdown_frequency(total_breakdowns, total_hours)
print()
print("Total breakdowns", total_breakdowns)
print("Total operating hours", total_hours)
print("Month breakdown frequency", round(month_frequency, 4))
