#!/usr/bin/env python3
"""Calculate every figure in the Maintenance Team KPI & Monthly Bonus Scheme, V2.

The daily sheet supplies operating hours, breakdown time, incident count, and
the recorded OEE. This module turns those rows into the monthly parameters,
the Year 1–3 score, and the bonus for each role.

Scoring follows the worked examples in the scheme:

* Higher is better (OEE, PM compliance, MTBF): score = actual / target x 100.
* Lower is better (breakdown frequency, breakdown hours, MTTR):
  score = target / actual x 100.
* A result above 100% is capped at 100% before it is rounded.
* Each score is rounded to one decimal, then multiplied by its weight.
* Weighted points are rounded to one decimal and added. The total is capped
  at 100%.

The spreadsheet header labels MTTR as 10%. The scheme scorecard weights MTTR
at 15% and MTBF at 10%. The scorecard is what this program uses.

PM compliance and lost-time injuries are not columns in the sheet. Pass them
in when you have them. Without them the program still calculates every
parameter it can, and reports the score range the missing lines can produce.
"""

from __future__ import annotations

import argparse
import csv
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET


ONE_DP = Decimal("0.1")
TWO_DP = Decimal("0.01")
HUNDRED = Decimal(100)
EXCEL_EPOCH = datetime(1899, 12, 30)
SHEET_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# Weights are the Year 1 scorecard shares. They sum to 1.
WEIGHTS = {
    "oee": Decimal("0.20"),
    "breakdown_frequency": Decimal("0.10"),
    "breakdown_hours": Decimal("0.15"),
    "pm_compliance": Decimal("0.20"),
    "mttr": Decimal("0.15"),
    "mtbf": Decimal("0.10"),
    "safety": Decimal("0.10"),
}

# Targets used in the score formula. "15 or fewer" is scored against 15:
# an actual equal to the target scores 100%.
YEAR_TARGETS = {
    1: {
        "oee": Decimal(82),
        "breakdown_frequency": Decimal(15),
        "breakdown_hours": Decimal(20),
        "pm_compliance": Decimal(98),
        "mttr": Decimal("2.5"),
        "mtbf": Decimal(150),
    },
    2: {
        "oee": Decimal(82),
        "breakdown_frequency": Decimal(10),
        "breakdown_hours": Decimal(12),
        "pm_compliance": Decimal(98),
        "mttr": Decimal(2),
        "mtbf": Decimal(500),
    },
    3: {
        "oee": Decimal(82),
        "breakdown_frequency": Decimal(10),
        "breakdown_hours": Decimal(5),
        "pm_compliance": Decimal(98),
        "mttr": Decimal("1.5"),
        "mtbf": Decimal(1000),
    },
}

# Monthly bonus opportunity, in pounds, from the two rate blocks in the scheme.
# The first block matches the day-shift worked examples (engineer 324,
# multi-skilled 486). The second block matches the night-shift examples
# (team leader 900, manager 1020).
BONUS_OPPORTUNITY = {
    ("Day", "Head of department"): Decimal(1235),
    ("Day", "Manager"): Decimal(918),
    ("Day", "Team leader"): Decimal(810),
    ("Day", "Multi-skilled"): Decimal(486),
    ("Day", "Mech / elec / maintenance / control / electrician"): Decimal(324),
    ("Day", "Handyman / electrician mate"): Decimal(243),
    ("Night", "Manager"): Decimal(1020),
    ("Night", "Team leader"): Decimal(900),
    ("Night", "Multi-skilled"): Decimal(540),
    ("Night", "Mech / elec / maintenance / control / electrician"): Decimal(360),
    ("Night", "Handyman / electrician mate"): Decimal(270),
}


def D(value: Decimal | int | str | float) -> Decimal:
    """Build a Decimal from a clean number. Floats go through str first."""
    if isinstance(value, Decimal):
        return value
    if isinstance(value, float):
        return Decimal(str(value))
    return Decimal(value)


def round_dp(value: Decimal, places: Decimal = ONE_DP) -> Decimal:
    return value.quantize(places, rounding=ROUND_HALF_UP)


def money(value: Decimal) -> str:
    amount = round_dp(value, TWO_DP)
    sign = "-" if amount < 0 else ""
    whole, frac = f"{abs(amount):.2f}".split(".")
    return f"{sign}£{int(whole):,}.{frac}"


# --- Component formulas from the scheme ---------------------------------


def availability_pct(actual_operating_time: Decimal, planned_production_time: Decimal) -> Decimal:
    """Availability = actual operating time / planned production time, in percent."""
    if planned_production_time == 0:
        raise ValueError("planned production time is 0")
    return D(actual_operating_time) / D(planned_production_time) * HUNDRED


def performance_pct(actual_production_rate: Decimal, ideal_production_rate: Decimal) -> Decimal:
    """Performance = actual production rate / ideal production rate, in percent."""
    if ideal_production_rate == 0:
        raise ValueError("ideal production rate is 0")
    return D(actual_production_rate) / D(ideal_production_rate) * HUNDRED


def quality_pct(good_units: Decimal, total_units: Decimal) -> Decimal:
    """Quality = good units / total units, in percent."""
    if total_units == 0:
        raise ValueError("total units is 0")
    return D(good_units) / D(total_units) * HUNDRED


def oee_pct(availability: Decimal, performance: Decimal, quality: Decimal) -> Decimal:
    """OEE = availability x performance x quality. Inputs and result are percents."""
    return D(availability) / HUNDRED * D(performance) / HUNDRED * D(quality) / HUNDRED * HUNDRED


def pm_compliance_pct(completed_pm: Decimal, planned_pm: Decimal) -> Decimal:
    """PM compliance = completed PM / planned PM, in percent."""
    if planned_pm == 0:
        raise ValueError("planned PM is 0")
    return D(completed_pm) / D(planned_pm) * HUNDRED


def mttr_hours(breakdown_hours: Decimal, incidents: Decimal) -> Decimal | None:
    """MTTR = breakdown hours / number of incidents. Undefined when there are no incidents."""
    if incidents == 0:
        return None
    return D(breakdown_hours) / D(incidents)


def mtbf_hours(operating_hours: Decimal, failures: Decimal) -> Decimal | None:
    """MTBF = operating hours / number of failures. Undefined when there are no failures."""
    if failures == 0:
        return None
    return D(operating_hours) / D(failures)


def breakdowns_per_operating_hour(incidents: Decimal, operating_hours: Decimal) -> Decimal | None:
    """Scheme sentence: breakdown frequency = breakdowns / operating hours."""
    if operating_hours == 0:
        return None
    return D(incidents) / D(operating_hours)


def score_higher_is_better(actual: Decimal, target: Decimal) -> Decimal:
    """actual / target x 100, capped at 100, rounded to one decimal."""
    actual, target = D(actual), D(target)
    if actual < 0 or target < 0:
        raise ValueError("actual and target must be zero or positive")
    if target == 0:
        raw = HUNDRED
    else:
        raw = actual / target * HUNDRED
    return _cap_and_round(raw)


def score_lower_is_better(actual: Decimal, target: Decimal) -> Decimal:
    """target / actual x 100, capped at 100, rounded to one decimal.

    An actual of zero is better than any positive target, so the score is 100.
    A target of zero is met only by an actual of zero.
    """
    actual, target = D(actual), D(target)
    if actual < 0 or target < 0:
        raise ValueError("actual and target must be zero or positive")
    if actual == 0:
        raw = HUNDRED
    elif target == 0:
        raw = Decimal(0)
    else:
        raw = target / actual * HUNDRED
    return _cap_and_round(raw)


def weighted_points(score: Decimal, weight: Decimal) -> Decimal:
    """score (0–100) x weight, rounded to one decimal. This is the line's contribution."""
    return round_dp(D(score) * D(weight))


def _cap_and_round(raw: Decimal) -> Decimal:
    if raw < 0:
        raw = Decimal(0)
    if raw > HUNDRED:
        raw = HUNDRED
    return round_dp(raw)


def safety_score(lti: Decimal) -> Decimal:
    """Zero lost-time injuries scores 100. Any LTI misses the safety KPI."""
    if D(lti) < 0:
        raise ValueError("LTI count cannot be negative")
    return Decimal("100.0") if D(lti) == 0 else Decimal("0.0")


# --- Daily rows ----------------------------------------------------------


@dataclass(frozen=True)
class DailyRecord:
    date: datetime | None
    availability: Decimal | None
    quality: Decimal | None
    operating_hours: Decimal
    breakdown_hours: Decimal
    incidents: Decimal
    oee: Decimal | None

    @property
    def breakdown_frequency(self) -> Decimal:
        """Daily incident count. The month is the sum of these counts."""
        return self.incidents

    @property
    def breakdowns_per_operating_hour(self) -> Decimal | None:
        return breakdowns_per_operating_hour(self.incidents, self.operating_hours)

    @property
    def mttr(self) -> Decimal | None:
        return mttr_hours(self.breakdown_hours, self.incidents)

    @property
    def mtbf(self) -> Decimal | None:
        return mtbf_hours(self.operating_hours, self.incidents)


@dataclass(frozen=True)
class DataWarning:
    code: str
    message: str


@dataclass(frozen=True)
class PeriodActuals:
    start: datetime | None
    end: datetime | None
    days: int
    operating_hours: Decimal
    breakdown_hours: Decimal
    incidents: Decimal
    oee: Decimal | None
    oee_weighted_by_hours: Decimal | None
    availability: Decimal | None
    availability_weighted_by_hours: Decimal | None
    quality: Decimal | None
    breakdowns_per_operating_hour: Decimal | None
    mttr: Decimal | None
    mtbf: Decimal | None
    warnings: tuple[DataWarning, ...] = ()


def summarise_period(rows: list[DailyRecord]) -> PeriodActuals:
    """Monthly parameters. Ratios use the period totals, which is what the scheme writes down."""
    if not rows:
        raise ValueError("no daily rows")

    operating = sum((row.operating_hours for row in rows), Decimal(0))
    breakdown = sum((row.breakdown_hours for row in rows), Decimal(0))
    incidents = sum((row.incidents for row in rows), Decimal(0))
    dates = [row.date for row in rows if row.date is not None]

    oee_rows = [row for row in rows if row.oee is not None]
    oee = sum((row.oee for row in oee_rows), Decimal(0)) / len(oee_rows) if oee_rows else None
    weighted_oee_hours = sum(
        (row.oee * row.operating_hours for row in oee_rows if row.operating_hours > 0),
        Decimal(0),
    )
    weighted_oee_den = sum(
        (row.operating_hours for row in oee_rows if row.operating_hours > 0),
        Decimal(0),
    )
    oee_weighted = weighted_oee_hours / weighted_oee_den if weighted_oee_den else None

    availability_rows = [row for row in rows if row.availability is not None]
    availability = (
        sum((row.availability for row in availability_rows), Decimal(0)) / len(availability_rows)
        if availability_rows
        else None
    )
    weighted_av_hours = sum(
        (row.availability * row.operating_hours for row in availability_rows if row.operating_hours > 0),
        Decimal(0),
    )
    weighted_av_den = sum(
        (row.operating_hours for row in availability_rows if row.operating_hours > 0),
        Decimal(0),
    )
    availability_weighted = weighted_av_hours / weighted_av_den if weighted_av_den else None

    quality_rows = [row for row in rows if row.quality is not None]
    quality = (
        sum((row.quality for row in quality_rows), Decimal(0)) / len(quality_rows)
        if quality_rows
        else None
    )

    warnings: list[DataWarning] = []
    for row in rows:
        if row.breakdown_hours > 0 and row.incidents == 0:
            when = row.date.date().isoformat() if row.date else "a row"
            warnings.append(
                DataWarning(
                    "breakdown_without_incident",
                    f"{when} has {row.breakdown_hours} breakdown hours and 0 incidents, "
                    "so that day's MTTR and MTBF cannot be calculated. "
                    "Those hours stay in the monthly breakdown-hour total.",
                )
            )
    if oee is not None and availability is not None and oee > availability:
        warnings.append(
            DataWarning(
                "oee_above_availability",
                f"Average OEE is {_show(oee, 1)}% and average availability is "
                f"{_show(availability, 1)}%. OEE is availability x performance x quality, "
                "so OEE cannot be higher than availability when all three are percentages "
                "of the same production time.",
            )
        )
    if quality is not None and oee is not None and quality < 50 and oee > 70:
        warnings.append(
            DataWarning(
                "quality_scale",
                f"Average quality is {_show(quality, 1)} while average OEE is {_show(oee, 1)}%. "
                "A quality rate of good units over total units in that range would not produce "
                "this OEE. The quality column is carried through as recorded.",
            )
        )

    return PeriodActuals(
        start=min(dates) if dates else None,
        end=max(dates) if dates else None,
        days=len(rows),
        operating_hours=operating,
        breakdown_hours=breakdown,
        incidents=incidents,
        oee=oee,
        oee_weighted_by_hours=oee_weighted,
        availability=availability,
        availability_weighted_by_hours=availability_weighted,
        quality=quality,
        breakdowns_per_operating_hour=breakdowns_per_operating_hour(incidents, operating),
        mttr=mttr_hours(breakdown, incidents),
        mtbf=mtbf_hours(operating, incidents),
        warnings=tuple(warnings),
    )


# --- Scorecard and bonus -------------------------------------------------


@dataclass(frozen=True)
class KpiLine:
    key: str
    name: str
    weight: Decimal
    higher_is_better: bool
    target: Decimal | None
    actual: Decimal | None
    score: Decimal | None
    points: Decimal | None
    formula: str


@dataclass(frozen=True)
class Gates:
    """Payout rules that sit on top of the KPI score."""

    lti: Decimal = Decimal(0)
    major_safety_violation: bool = False
    major_epw: bool = False
    dishonest_reporting: bool = False
    red_audit: bool = False
    green_audit_amount: Decimal = Decimal(0)

    def __post_init__(self) -> None:
        if self.red_audit and self.green_audit_amount != 0:
            raise ValueError("a red audit and a green audit are different outcomes")
        if self.lti < 0 or self.green_audit_amount < 0:
            raise ValueError("LTI count and green-audit amount cannot be negative")


@dataclass(frozen=True)
class BonusLine:
    shift: str
    role: str
    opportunity: Decimal
    before_gates: Decimal
    paid: Decimal
    gate_notes: tuple[str, ...]


@dataclass
class Scorecard:
    year: int
    lines: list[KpiLine] = field(default_factory=list)
    points: Decimal = Decimal(0)
    open_points: Decimal = Decimal(0)
    final_score: Decimal | None = None
    minimum_score: Decimal = Decimal(0)
    maximum_score: Decimal = Decimal(0)
    bonuses: list[BonusLine] = field(default_factory=list)
    gate_notes: tuple[str, ...] = ()

    @property
    def complete(self) -> bool:
        return self.final_score is not None


def calculate_scorecard(
    actuals: PeriodActuals,
    *,
    year: int = 1,
    pm_compliance: Decimal | None = None,
    lti: Decimal | None = None,
    gates: Gates | None = None,
) -> Scorecard:
    """Score the period against one target year and price every role.

    `pm_compliance` is a percent. `lti` is the count of lost-time injuries.
    Leave either as None when the sheet does not contain it. The final score
    and the bonus amounts are filled in only when every KPI has an actual.
    """
    if year not in YEAR_TARGETS:
        raise ValueError(f"year must be one of {sorted(YEAR_TARGETS)}")
    targets = YEAR_TARGETS[year]
    applied_gates = gates if gates is not None else Gates(lti=lti or 0)

    if lti is not None and applied_gates.lti != D(lti):
        raise ValueError("pass the LTI count once, either as lti or inside gates")
    # The safety line is scored only when the caller passes an LTI count.
    # Gates.lti defaults to 0, so a gates object on its own must not award
    # the safety points.
    safety_actual = None if lti is None else D(lti)

    specs = [
        (
            "oee",
            "OEE",
            True,
            targets["oee"],
            actuals.oee,
            "Recorded daily OEE, averaged across the days in the file. "
            "Score = actual / target x 100.",
        ),
        (
            "breakdown_frequency",
            "Breakdown frequency",
            False,
            targets["breakdown_frequency"],
            actuals.incidents,
            "Count of incidents in the period. The year target is a count per "
            "line per month. Score = target / actual x 100. "
            "Breakdowns per operating hour is the rate form of the same count.",
        ),
        (
            "breakdown_hours",
            "Breakdown hours",
            False,
            targets["breakdown_hours"],
            actuals.breakdown_hours,
            "Sum of breakdown hours. Score = target / actual x 100.",
        ),
        (
            "pm_compliance",
            "PM compliance",
            True,
            targets["pm_compliance"],
            None if pm_compliance is None else D(pm_compliance),
            "Completed PM / planned PM, in percent. Score = actual / target x 100.",
        ),
        (
            "mttr",
            "MTTR",
            False,
            targets["mttr"],
            actuals.mttr if actuals.incidents > 0 else (Decimal(0) if actuals.breakdown_hours == 0 else None),
            "Total breakdown hours / number of incidents. Score = target / actual x 100. "
            "With incidents and breakdown hours both zero, MTTR scores 100.",
        ),
        (
            "mtbf",
            "MTBF",
            True,
            targets["mtbf"],
            actuals.mtbf,
            "Total operating hours / number of failures. Score = actual / target x 100. "
            "Zero failures scores 100.",
        ),
        (
            "safety",
            "Safety",
            True,
            Decimal(0),
            safety_actual,
            "Zero lost-time injuries scores 100. Any LTI scores 0 on this line. "
            "An LTI also halves the bonus after the score is calculated.",
        ),
    ]

    card = Scorecard(year=year)
    for key, name, higher, target, actual, formula in specs:
        weight = WEIGHTS[key]
        if key == "mtbf" and actuals.incidents == 0 and actuals.operating_hours > 0:
            score = Decimal("100.0")
            points = weighted_points(score, weight)
            actual_for_line = None
        elif key == "safety":
            if actual is None:
                score = None
                points = None
            else:
                score = safety_score(actual)
                points = weighted_points(score, weight)
            actual_for_line = actual
        elif actual is None:
            score = None
            points = None
            actual_for_line = None
        elif higher:
            score = score_higher_is_better(actual, target)
            points = weighted_points(score, weight)
            actual_for_line = actual
        else:
            score = score_lower_is_better(actual, target)
            points = weighted_points(score, weight)
            actual_for_line = actual

        card.lines.append(
            KpiLine(
                key=key,
                name=name,
                weight=weight,
                higher_is_better=higher,
                target=target,
                actual=actual_for_line,
                score=score,
                points=points,
                formula=formula,
            )
        )
        if points is None:
            card.open_points += round_dp(weight * HUNDRED)
        else:
            card.points += points

    card.minimum_score = min(HUNDRED, card.points)
    card.maximum_score = min(HUNDRED, card.points + card.open_points)
    if card.open_points == 0:
        card.final_score = min(HUNDRED, card.points)
        card.bonuses, card.gate_notes = price_bonus(card.final_score, applied_gates)
    return card


def price_bonus(final_score: Decimal, gates: Gates) -> tuple[list[BonusLine], tuple[str, ...]]:
    """Bonus paid = opportunity x final score. Then the gatekeeper rules."""
    notes: list[str] = []
    stopped = False
    if gates.major_safety_violation:
        notes.append("Major safety violation: bonus stopped.")
        stopped = True
    if gates.major_epw:
        notes.append("Major EPW: bonus stopped.")
        stopped = True
    if gates.dishonest_reporting:
        notes.append("Misleading or dishonest reporting: bonus stopped.")
        stopped = True
    multiplier = Decimal(1)
    if not stopped and gates.lti > 0:
        multiplier *= Decimal("0.5")
        notes.append("LTI is above zero: total bonus reduced by 50%.")
    if not stopped and gates.red_audit:
        multiplier *= Decimal("0.5")
        notes.append("Red audit: bonus reduced to 50%.")
    extra = Decimal(0) if stopped else gates.green_audit_amount
    if extra:
        notes.append(f"Green audit: {money(extra)} added on top of the bonus.")

    lines: list[BonusLine] = []
    for (shift, role), opportunity in BONUS_OPPORTUNITY.items():
        before = opportunity * D(final_score) / HUNDRED
        paid = Decimal(0) if stopped else before * multiplier + extra
        lines.append(
            BonusLine(
                shift=shift,
                role=role,
                opportunity=opportunity,
                before_gates=round_dp(before, TWO_DP),
                paid=round_dp(paid, TWO_DP),
                gate_notes=tuple(notes),
            )
        )
    return lines, tuple(notes)


# --- Workbook ------------------------------------------------------------


def load_daily_rows(path: str | Path) -> list[DailyRecord]:
    """Read the daily sheet. Header names are matched loosely."""
    table = _read_xlsx_table(Path(path))
    if not table:
        raise ValueError(f"{path} has no rows")
    headers = [_canonical_header(cell) for cell in table[0]]
    columns = _column_index(headers)
    required = ("operating_hours", "breakdown_hours", "incidents")
    missing = [name for name in required if name not in columns]
    if missing:
        raise ValueError(f"{path} is missing columns: {', '.join(missing)}")

    rows: list[DailyRecord] = []
    for raw in table[1:]:
        if _row_is_blank(raw):
            continue
        rows.append(
            DailyRecord(
                date=_cell_date(raw, columns.get("date")),
                availability=_cell_decimal(raw, columns.get("availability")),
                quality=_cell_decimal(raw, columns.get("quality")),
                operating_hours=_cell_decimal(raw, columns.get("operating_hours")) or Decimal(0),
                breakdown_hours=_cell_decimal(raw, columns.get("breakdown_hours")) or Decimal(0),
                incidents=_cell_decimal(raw, columns.get("incidents")) or Decimal(0),
                oee=_cell_decimal(raw, columns.get("oee")),
            )
        )
    if not rows:
        raise ValueError(f"{path} has headers but no daily rows")
    return rows


def write_daily_csv(path: str | Path, rows: list[DailyRecord]) -> None:
    """Daily parameters, one row per day, including the columns the sheet left blank."""
    with Path(path).open("w", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "date",
                "availability",
                "quality",
                "operating_hours",
                "breakdown_hours",
                "incidents",
                "oee",
                "breakdown_frequency",
                "breakdowns_per_operating_hour",
                "mttr",
                "mtbf",
            ]
        )
        for row in rows:
            writer.writerow(
                [
                    row.date.date().isoformat() if row.date else "",
                    _csv(row.availability),
                    _csv(row.quality),
                    _csv(row.operating_hours),
                    _csv(row.breakdown_hours),
                    _csv(row.incidents),
                    _csv(row.oee),
                    _csv(row.breakdown_frequency),
                    _csv(row.breakdowns_per_operating_hour),
                    _csv(row.mttr),
                    _csv(row.mtbf),
                ]
            )


def _csv(value: Decimal | None) -> str:
    if value is None:
        return ""
    return format(value, "f")


def _read_xlsx_table(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as workbook:
        shared = _shared_strings(workbook)
        sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
    rows: list[list[str]] = []
    for row in sheet.findall("m:sheetData/m:row", SHEET_NS):
        values: dict[int, str] = {}
        for cell in row.findall("m:c", SHEET_NS):
            ref = cell.get("r") or "A1"
            column = _column_number(ref)
            values[column] = _cell_text(cell, shared)
        if not values:
            rows.append([])
            continue
        width = max(values)
        rows.append([values.get(index, "") for index in range(1, width + 1)])
    return rows


def _shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("m:si", SHEET_NS):
        strings.append(
            "".join(
                node.text or ""
                for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
            )
        )
    return strings


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    kind = cell.get("t")
    value = cell.find("m:v", SHEET_NS)
    if kind == "s" and value is not None and value.text:
        return shared[int(value.text)]
    if kind == "inlineStr":
        inline = cell.find("m:is", SHEET_NS)
        if inline is None:
            return ""
        return "".join(
            node.text or ""
            for node in inline.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        )
    if value is None or value.text is None:
        return ""
    return value.text


def _column_number(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    number = 0
    for char in letters:
        number = number * 26 + (ord(char.upper()) - 64)
    return number


def _canonical_header(label: str) -> str:
    compact = "".join(char for char in label.lower() if char.isalnum())
    aliases = {
        "date": "date",
        "availability": "availability",
        "quality": "quality",
        "operatinghours": "operating_hours",
        "breakdowntime": "breakdown_hours",
        "breakdownhours": "breakdown_hours_calculated",
        "noofincidentsnooffailures": "incidents",
        "noofincidents": "incidents",
        "incidents": "incidents",
        "failures": "incidents",
        "nooffailures": "incidents",
        "oee": "oee",
    }
    return aliases.get(compact, compact)


def _column_index(headers: list[str]) -> dict[str, int]:
    found: dict[str, int] = {}
    for index, header in enumerate(headers):
        if header in {
            "date",
            "availability",
            "quality",
            "operating_hours",
            "breakdown_hours",
            "incidents",
            "oee",
        }:
            found.setdefault(header, index)
    return found


def _row_is_blank(raw: list[str]) -> bool:
    return all(not cell.strip() for cell in raw)


def _cell_decimal(raw: list[str], index: int | None) -> Decimal | None:
    if index is None or index >= len(raw):
        return None
    text = raw[index].strip()
    if not text:
        return None
    return Decimal(text)


def _cell_date(raw: list[str], index: int | None) -> datetime | None:
    if index is None or index >= len(raw):
        return None
    text = raw[index].strip()
    if not text:
        return None
    if text.isdigit():
        return EXCEL_EPOCH + timedelta(days=int(text))
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    raise ValueError(f"unrecognised date: {text}")


# --- Report --------------------------------------------------------------


def render_report(source: str, rows: list[DailyRecord], actuals: PeriodActuals, card: Scorecard) -> str:
    lines: list[str] = []
    period = _period_label(actuals)
    lines.append("Maintenance KPI calculation")
    lines.append(f"Source: {source}")
    lines.append(f"Period: {period}")
    lines.append(f"Targets: year {card.year}")
    lines.append("")
    lines.append("Parameters")
    lines.append(f"  Days                              {actuals.days}")
    lines.append(f"  Operating hours                   {_show(actuals.operating_hours, 0)}")
    lines.append(f"  Breakdown hours                   {_show(actuals.breakdown_hours, 0)}")
    lines.append(f"  Incidents (breakdown frequency)   {_show(actuals.incidents, 0)}")
    lines.append(
        f"  Breakdowns per operating hour     {_show(actuals.breakdowns_per_operating_hour, 4)}"
    )
    lines.append(f"  OEE, daily average                {_pct(actuals.oee)}")
    lines.append(f"  OEE, weighted by operating hours  {_pct(actuals.oee_weighted_by_hours)}")
    lines.append(f"  Availability, daily average       {_pct(actuals.availability)}")
    lines.append(f"  Availability, hour-weighted       {_pct(actuals.availability_weighted_by_hours)}")
    lines.append(f"  Quality, daily average            {_show(actuals.quality, 1)}")
    lines.append(f"  MTTR, hours                       {_show(actuals.mttr, 2)}")
    lines.append(f"  MTBF, hours                       {_show(actuals.mtbf, 1)}")
    lines.append("")
    lines.append("Scorecard")
    lines.append(
        f"  {'KPI':<24} {'Weight':>8} {'Target':>10} {'Actual':>10} {'Score':>8} {'Points':>8}"
    )
    for line in card.lines:
        lines.append(
            f"  {line.name:<24} {_weight(line.weight):>8} {_target(line):>10} "
            f"{_actual(line):>10} {_score(line.score):>8} {_score(line.points):>8}"
        )
    measured_possible = round_dp(HUNDRED - card.open_points)
    lines.append(f"  {'Measured points':<24} {card.points} of {measured_possible}")
    lines.append("  Scores use the exact totals, then round to one decimal.")
    if card.final_score is None:
        lines.append(f"  Open points                   {card.open_points}")
        lines.append(f"  Score if the open KPIs score 0     {card.minimum_score}%")
        lines.append(f"  Score if the open KPIs score 100   {card.maximum_score}%")
    else:
        lines.append(f"  Final KPI score               {card.final_score}%")
    lines.append("")
    lines.append("How each line is calculated")
    for line in card.lines:
        lines.append(f"  {line.name}: {line.formula}")
    if actuals.warnings:
        lines.append("")
        lines.append("Checks")
        for warning in actuals.warnings:
            lines.append(f"  {warning.message}")
    if card.final_score is not None:
        lines.append("")
        lines.append("Bonus")
        lines.append(f"  Final KPI score {card.final_score}%")
        for note in card.gate_notes:
            lines.append(f"  {note}")
        lines.append(
            f"  {'Shift':<8} {'Role':<52} {'Opportunity':>14} {'Before gates':>14} {'Paid':>14}"
        )
        for bonus in card.bonuses:
            lines.append(
                f"  {bonus.shift:<8} {bonus.role:<52} {money(bonus.opportunity):>14} "
                f"{money(bonus.before_gates):>14} {money(bonus.paid):>14}"
            )
    else:
        lines.append("")
        lines.append("Bonus")
        lines.append("  Not priced. Pass --pm and --lti once those results are known.")
    return "\n".join(lines)


def _period_label(actuals: PeriodActuals) -> str:
    if actuals.start and actuals.end:
        return (
            f"{actuals.start.date().isoformat()} to {actuals.end.date().isoformat()} "
            f"({actuals.days} days)"
        )
    return f"{actuals.days} days"


def _show(value: Decimal | None, places: int) -> str:
    if value is None:
        return "—"
    quantum = Decimal(1).scaleb(-places)
    return format(round_dp(value, quantum), "f")


def _pct(value: Decimal | None) -> str:
    if value is None:
        return "—"
    return f"{_show(value, 1)}%"


def _weight(weight: Decimal) -> str:
    return f"{round_dp(weight * HUNDRED, Decimal(1))}%"


def _target(line: KpiLine) -> str:
    if line.key == "safety":
        return "0 LTI"
    if line.target is None:
        return "—"
    if line.key in {"oee", "pm_compliance"}:
        return f"{_show(line.target, 1)}%"
    if line.key == "breakdown_frequency":
        return _show(line.target, 0)
    return _show(line.target, 1)


def _actual(line: KpiLine) -> str:
    if line.actual is None:
        if line.key == "mtbf" and line.score == Decimal("100.0"):
            return "no failures"
        return "—"
    if line.key in {"oee", "pm_compliance"}:
        return f"{_show(line.actual, 1)}%"
    if line.key == "safety":
        return _show(line.actual, 0)
    if line.key == "breakdown_frequency":
        return _show(line.actual, 0)
    if line.key == "mttr":
        return _show(line.actual, 2)
    return _show(line.actual, 1)


def _score(value: Decimal | None) -> str:
    if value is None:
        return "—"
    return format(value, "f")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", default="ref.xlsx", help="daily workbook, default ref.xlsx")
    parser.add_argument("--year", type=int, default=1, choices=sorted(YEAR_TARGETS))
    parser.add_argument("--pm", help="PM compliance in percent, for example 96")
    parser.add_argument("--pm-completed", help="completed PM tasks, used with --pm-planned")
    parser.add_argument("--pm-planned", help="planned PM tasks, used with --pm-completed")
    parser.add_argument("--lti", help="lost-time injury count. 0 scores the safety line at 100")
    parser.add_argument("--major-safety", action="store_true", help="stop the bonus")
    parser.add_argument("--major-epw", action="store_true", help="stop the bonus")
    parser.add_argument("--dishonest-reporting", action="store_true", help="stop the bonus")
    parser.add_argument("--red-audit", action="store_true", help="cut the bonus by 50 percent")
    parser.add_argument("--green-audit", default="0", help="pounds added on top for a green audit")
    parser.add_argument("--csv", help="write one row per day to this CSV path")
    args = parser.parse_args(argv)

    pm = _optional_decimal(args.pm, "--pm")
    if args.pm_completed or args.pm_planned:
        if args.pm:
            parser.error("pass either --pm or both --pm-completed and --pm-planned")
        if not args.pm_completed or not args.pm_planned:
            parser.error("pass both --pm-completed and --pm-planned")
        pm = pm_compliance_pct(D(args.pm_completed), D(args.pm_planned))
    lti = _optional_decimal(args.lti, "--lti")
    gates = Gates(
        lti=lti or 0,
        major_safety_violation=args.major_safety,
        major_epw=args.major_epw,
        dishonest_reporting=args.dishonest_reporting,
        red_audit=args.red_audit,
        green_audit_amount=D(args.green_audit),
    )

    rows = load_daily_rows(args.workbook)
    actuals = summarise_period(rows)
    card = calculate_scorecard(
        actuals,
        year=args.year,
        pm_compliance=pm,
        lti=lti,
        gates=gates if lti is not None else None,
    )
    if args.csv:
        write_daily_csv(args.csv, rows)
    print(render_report(args.workbook, rows, actuals, card))
    return 0


def _optional_decimal(text: str | None, flag: str) -> Decimal | None:
    if text is None:
        return None
    try:
        return D(text)
    except Exception as error:
        raise SystemExit(f"{flag} must be a number") from error


if __name__ == "__main__":
    sys.exit(main())
