# OEE = Availability x Performance x Quality
# Type percentages. 80 means 80%.
# Click Run Python File in VS Code.

performance = 95


def oee(availability, performance, quality):
    return (availability / 100) * (performance / 100) * (quality / 100) * 100


# availability, quality  (from the daily sheet)
rows = [
    (70.2, 15.7),
    (72, 17),
    (69, 18),
    (77, 15),
    (66, 16),
    (55, 19),
    (69, 14),
    (78, 15),
    (67, 14),
    (80, 10),
    (78, 16),
    (70, 18),
    (69, 20),
    (69, 20),
    (75, 10),
    (77, 17),
    (80, 16),
    (67, 20),
    (78, 20),
]

print("Performance used:", performance)
print("Day   Availability   Quality   OEE")

total = 0
for day, (availability, quality) in enumerate(rows, start=1):
    result = oee(availability, performance, quality)
    total += result
    print(day, availability, quality, round(result, 2))

print("Average OEE", round(total / len(rows), 2))
