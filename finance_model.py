"""Transparent, deterministic calculations for the HK Real Estate BI demo.

All money amounts are HKD. These calculations are *illustrative*, not a
property-development cash-flow model or a lender covenant test.
"""

from dataclasses import dataclass, replace
from math import isfinite
from typing import Iterable


SQFT_PER_SQM = 10.7639


@dataclass(frozen=True)
class ProjectInputs:
    gfa_sqm: float
    land_cost_hkd: float
    construction_cost_per_gfa_sqft: float
    selling_price_per_saleable_sqft: float
    tenure_years: float
    interest_rate_pct: float
    ltv_pct: float
    saleable_efficiency_pct: float = 100.0
    marketing_rate_pct: float = 5.0
    average_debt_utilization_pct: float = 100.0


@dataclass(frozen=True)
class ProjectResults:
    gross_sqft: float
    saleable_sqft: float
    revenue_hkd: float
    land_hkd: float
    construction_hkd: float
    loan_principal_hkd: float
    interest_hkd: float
    marketing_hkd: float
    total_cost_hkd: float
    profit_hkd: float
    roi_on_cost: float
    break_even_per_saleable_sqft: float


def _validate(inputs: ProjectInputs) -> None:
    values = vars(inputs)
    if any(not isfinite(float(v)) for v in values.values()):
        raise ValueError("所有輸入必須為有限數值。")
    if inputs.gfa_sqm <= 0 or inputs.tenure_years <= 0:
        raise ValueError("總樓面面積及開發年期必須大於零。")
    if any(
        v < 0
        for v in (
            inputs.land_cost_hkd,
            inputs.construction_cost_per_gfa_sqft,
            inputs.selling_price_per_saleable_sqft,
            inputs.interest_rate_pct,
        )
    ):
        raise ValueError("金額、售價及利率不可為負數。")
    for label, value in (
        ("LTV", inputs.ltv_pct),
        ("平均貸款使用比例", inputs.average_debt_utilization_pct),
    ):
        if not 0 <= value <= 100:
            raise ValueError(f"{label} 必須介乎 0–100%。")
    if not 0 < inputs.saleable_efficiency_pct <= 100:
        raise ValueError("實用面積比例必須大於 0%，且不超過 100%。")
    if not 0 <= inputs.marketing_rate_pct < 100:
        raise ValueError("營銷費率必須介乎 0% 至低於 100%。")


def calculate_project(inputs: ProjectInputs) -> ProjectResults:
    """Compute a cumulative *return on cost*, with explicit financing assumptions.

    Financing is SIMPLE interest on an assumed average outstanding loan:
       (land + construction) × LTV × utilization × annual rate × years.
    This is NOT an amortization, cash-flow, IRR, or equity-return calculation.

    Marketing is variable: marketing_rate × *actual* sale proceeds.
    Hence break-even solves P × area × (1 - marketing_rate) = fixed costs.
    """
    _validate(inputs)
    gross_sqft = inputs.gfa_sqm * SQFT_PER_SQM
    saleable_sqft = gross_sqft * inputs.saleable_efficiency_pct / 100.0
    construction = gross_sqft * inputs.construction_cost_per_gfa_sqft
    fixed_development_cost = inputs.land_cost_hkd + construction
    if fixed_development_cost <= 0:
        raise ValueError("土地及建造成本不可同時為零，否則 ROI 分母沒有意義。")

    principal = fixed_development_cost * inputs.ltv_pct / 100.0
    interest = (
        principal
        * inputs.average_debt_utilization_pct / 100.0
        * inputs.interest_rate_pct / 100.0
        * inputs.tenure_years
    )
    revenue = saleable_sqft * inputs.selling_price_per_saleable_sqft
    marketing = revenue * inputs.marketing_rate_pct / 100.0
    fixed_costs = fixed_development_cost + interest
    total_cost = fixed_costs + marketing
    profit = revenue - total_cost
    break_even = fixed_costs / (
        saleable_sqft * (1 - inputs.marketing_rate_pct / 100.0)
    )
    return ProjectResults(
        gross_sqft=gross_sqft,
        saleable_sqft=saleable_sqft,
        revenue_hkd=revenue,
        land_hkd=inputs.land_cost_hkd,
        construction_hkd=construction,
        loan_principal_hkd=principal,
        interest_hkd=interest,
        marketing_hkd=marketing,
        total_cost_hkd=total_cost,
        profit_hkd=profit,
        roi_on_cost=profit / total_cost,
        break_even_per_saleable_sqft=break_even,
    )


def sensitivity_grid(
    inputs: ProjectInputs,
    price_shocks_pct: Iterable[float],
    interest_rates_pct: Iterable[float],
) -> list[list[float]]:
    """Calculate each grid cell with the SAME model as the ROI calculator."""
    rows = []
    for shock in price_shocks_pct:
        if shock < -100:
            raise ValueError("樓價下行情景不能低於 -100%。")
        row = []
        for rate in interest_rates_pct:
            scenario = replace(
                inputs,
                selling_price_per_saleable_sqft=(
                    inputs.selling_price_per_saleable_sqft * (1 + shock / 100.0)
                ),
                interest_rate_pct=rate,
            )
            row.append(calculate_project(scenario).roi_on_cost)
        rows.append(row)
    return rows
