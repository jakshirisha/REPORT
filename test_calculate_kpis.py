"""The scorecard examples in the scheme, and the September sheet."""

import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from calculate_metrics import breakdown_frequency, calculate_sheet, mtbf, mttr, oee
from calculate_kpis import (
    DailyRecord,
    Gates,
    availability_pct,
    calculate_scorecard,
    load_daily_rows,
    mtbf_hours,
    mttr_hours,
    oee_pct,
    performance_pct,
    pm_compliance_pct,
    price_bonus,
    quality_pct,
    score_higher_is_better,
    score_lower_is_better,
    summarise_period,
    weighted_points,
)


ROOT = Path(__file__).resolve().parent


def period_with(**overrides):
    rows = [
        DailyRecord(
            date=datetime(2026, 9, 1),
            availability=Decimal(80),
            quality=Decimal(99),
            operating_hours=Decimal(100),
            breakdown_hours=Decimal(0),
            incidents=Decimal(0),
            oee=Decimal(82),
        )
    ]
    actuals = summarise_period(rows)
    for name, value in overrides.items():
        object.__setattr__(actuals, name, value)
    return actuals


class SchemeExamples(unittest.TestCase):
    def test_component_formulas(self):
        self.assertEqual(availability_pct(20, 25), Decimal(80))
        self.assertEqual(performance_pct(90, 100), Decimal(90))
        self.assertEqual(quality_pct(95, 100), Decimal(95))
        self.assertEqual(oee_pct(80, 90, 95), Decimal("68.4"))
        self.assertEqual(pm_compliance_pct(96, 100), Decimal(96))
        self.assertEqual(mttr_hours(Decimal(3), Decimal(1)), Decimal(3))
        self.assertIsNone(mttr_hours(Decimal(0), Decimal(0)))
        self.assertEqual(mtbf_hours(Decimal(160), Decimal(1)), Decimal(160))

    def test_worked_example_lines(self):
        self.assertEqual(score_higher_is_better(78, 82), Decimal("95.1"))
        self.assertEqual(weighted_points(Decimal("95.1"), Decimal("0.20")), Decimal("19.0"))
        self.assertEqual(score_lower_is_better(9, 10), Decimal("100.0"))
        self.assertEqual(weighted_points(Decimal("100.0"), Decimal("0.10")), Decimal("10.0"))
        self.assertEqual(score_lower_is_better(18, 20), Decimal("100.0"))
        self.assertEqual(weighted_points(Decimal("100"), Decimal("0.15")), Decimal("15.0"))
        self.assertEqual(score_higher_is_better(96, 98), Decimal("98.0"))
        self.assertEqual(weighted_points(Decimal("98.0"), Decimal("0.20")), Decimal("19.6"))
        self.assertEqual(score_lower_is_better(Decimal("3.0"), Decimal("2.5")), Decimal("83.3"))
        self.assertEqual(weighted_points(Decimal("83.3"), Decimal("0.15")), Decimal("12.5"))
        self.assertEqual(score_higher_is_better(160, 150), Decimal("100.0"))
        self.assertEqual(weighted_points(Decimal("100"), Decimal("0.10")), Decimal("10.0"))

    def test_worked_example_totals_and_bonus(self):
        actuals = period_with(
            incidents=Decimal(9),
            breakdown_hours=Decimal(18),
            operating_hours=Decimal(1440),
            oee=Decimal(78),
            mttr=Decimal(3),
            mtbf=Decimal(160),
        )
        card = calculate_scorecard(actuals, year=1, pm_compliance=Decimal(96), lti=Decimal(0))
        self.assertEqual(card.final_score, Decimal("96.1"))
        by_role = {(line.shift, line.role): line for line in card.bonuses}
        self.assertEqual(
            by_role[("Day", "Mech / elec / maintenance / control / electrician")].paid,
            Decimal("311.36"),
        )
        self.assertEqual(by_role[("Night", "Manager")].paid, Decimal("980.22"))
        self.assertEqual(by_role[("Night", "Team leader")].paid, Decimal("864.90"))
        # 486 x 96.1% = 467.046, which is 467.05 to the penny.
        self.assertEqual(by_role[("Day", "Multi-skilled")].paid, Decimal("467.05"))

    def test_gates(self):
        before, notes = price_bonus(Decimal("100"), Gates(lti=Decimal(1)))
        self.assertEqual(before[0].paid, round(before[0].opportunity * Decimal("0.5"), 2))
        self.assertTrue(any("50%" in note for note in notes))

        stopped, stop_notes = price_bonus(
            Decimal("100"),
            Gates(major_safety_violation=True, green_audit_amount=Decimal(50)),
        )
        self.assertEqual(stopped[0].paid, Decimal("0.00"))
        self.assertTrue(stop_notes)

        halved, _ = price_bonus(Decimal("100"), Gates(lti=Decimal(1), red_audit=True))
        self.assertEqual(halved[0].paid, round(halved[0].opportunity * Decimal("0.25"), 2))

        with self.assertRaises(ValueError):
            Gates(red_audit=True, green_audit_amount=Decimal(10))


class SeptemberSheet(unittest.TestCase):
    def test_file_parameters_and_year_1_score(self):
        rows = load_daily_rows(ROOT / "ref.xlsx")
        actuals = summarise_period(rows)
        self.assertEqual(actuals.days, 19)
        self.assertEqual(actuals.start, datetime(2026, 9, 1))
        self.assertEqual(actuals.end, datetime(2026, 9, 19))
        self.assertEqual(actuals.operating_hours, Decimal(686))
        self.assertEqual(actuals.breakdown_hours, Decimal(67))
        self.assertEqual(actuals.incidents, Decimal(37))
        self.assertEqual(actuals.mttr, Decimal(67) / Decimal(37))
        self.assertEqual(actuals.mtbf, Decimal(686) / Decimal(37))

        self.assertEqual(rows[0].mttr, Decimal("2.5"))
        self.assertEqual(rows[0].mtbf, Decimal("12.5"))
        self.assertIsNone(rows[9].mttr)
        self.assertEqual(rows[18].breakdown_hours, Decimal(2))
        self.assertEqual(rows[18].incidents, Decimal(0))

        card = calculate_scorecard(actuals, year=1)
        by_key = {line.key: line for line in card.lines}
        self.assertEqual(by_key["oee"].score, Decimal("100.0"))
        self.assertEqual(by_key["oee"].points, Decimal("20.0"))
        self.assertEqual(by_key["breakdown_frequency"].score, Decimal("40.5"))
        self.assertEqual(by_key["breakdown_frequency"].points, Decimal("4.1"))
        self.assertEqual(by_key["breakdown_hours"].score, Decimal("29.9"))
        self.assertEqual(by_key["breakdown_hours"].points, Decimal("4.5"))
        self.assertIsNone(by_key["pm_compliance"].score)
        self.assertEqual(by_key["mttr"].score, Decimal("100.0"))
        self.assertEqual(by_key["mttr"].points, Decimal("15.0"))
        self.assertEqual(by_key["mtbf"].score, Decimal("12.4"))
        self.assertEqual(by_key["mtbf"].points, Decimal("1.2"))
        self.assertIsNone(by_key["safety"].score)
        self.assertEqual(card.points, Decimal("44.8"))
        self.assertEqual(card.open_points, Decimal(30))
        self.assertEqual(card.maximum_score, Decimal("74.8"))
        self.assertIsNone(card.final_score)

        codes = {warning.code for warning in actuals.warnings}
        self.assertIn("breakdown_without_incident", codes)
        self.assertIn("oee_above_availability", codes)

    def test_full_score_when_pm_and_safety_are_supplied(self):
        actuals = summarise_period(load_daily_rows(ROOT / "ref.xlsx"))
        card = calculate_scorecard(
            actuals,
            year=1,
            pm_compliance=Decimal(100),
            lti=Decimal(0),
        )
        self.assertEqual(card.final_score, Decimal("74.8"))
        engineer = next(
            line
            for line in card.bonuses
            if line.shift == "Day" and line.role.startswith("Mech")
        )
        self.assertEqual(engineer.opportunity, Decimal(324))
        self.assertEqual(engineer.paid, Decimal("242.35"))


class FourMetrics(unittest.TestCase):
    def test_formulas(self):
        self.assertAlmostEqual(oee(80, 90, 95), 68.4)
        self.assertAlmostEqual(breakdown_frequency(2, 25), 0.08)
        self.assertAlmostEqual(mttr(5, 2), 2.5)
        self.assertAlmostEqual(mtbf(25, 2), 12.5)

    def test_september_totals(self):
        _daily, month = calculate_sheet(ROOT / "ref.xlsx")
        self.assertAlmostEqual(month["oee"], 1676.9 / 19)
        self.assertAlmostEqual(month["breakdown_frequency"], 37 / 686)
        self.assertEqual(month["breakdowns"], 37)
        self.assertAlmostEqual(month["mttr"], 67 / 37)
        self.assertAlmostEqual(month["mtbf"], 686 / 37)


if __name__ == "__main__":
    unittest.main()
