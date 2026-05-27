"""
Tax rate calculation engine for Coretax BP21, BPMP, and BPPU.
Ported from Excel VBA formulas in the official Coretax templates.
BPA1 does not need rate calculation (PPh computed server-side).
"""

from modules.reference_data import (
    BP21_TAX_OBJECT_LOOKUP,
    BPMP_TAX_OBJECT_LOOKUP,
    BPPU_TAX_OBJECT_LOOKUP,
    PTKP_CATEGORY_A,
    PTKP_CATEGORY_B,
    PTKP_CATEGORY_C,
)


# =============================================================================
# TER (Tarif Efektif Rata-rata) Lookup Tables
# Format: list of (upper_bound, rate)
# If income <= upper_bound, use that rate.
# Last entry has no upper bound (uses float('inf')).
# Shared by BP21 and BPMP.
# =============================================================================

TER_CATEGORY_A = [
    (5_400_000, 0),
    (5_650_000, 0.25),
    (5_950_000, 0.5),
    (6_300_000, 0.75),
    (6_750_000, 1),
    (7_500_000, 1.25),
    (8_550_000, 1.5),
    (9_650_000, 1.75),
    (10_050_000, 2),
    (10_350_000, 2.25),
    (10_700_000, 2.5),
    (11_050_000, 3),
    (11_600_000, 3.5),
    (12_500_000, 4),
    (13_750_000, 5),
    (15_100_000, 6),
    (16_950_000, 7),
    (19_750_000, 8),
    (24_150_000, 9),
    (26_450_000, 10),
    (28_000_000, 11),
    (30_050_000, 12),
    (32_400_000, 13),
    (35_400_000, 14),
    (39_100_000, 15),
    (43_850_000, 16),
    (47_800_000, 17),
    (51_400_000, 18),
    (56_300_000, 19),
    (62_200_000, 20),
    (68_600_000, 21),
    (77_500_000, 22),
    (89_000_000, 23),
    (103_000_000, 24),
    (125_000_000, 25),
    (157_000_000, 26),
    (206_000_000, 27),
    (337_000_000, 28),
    (454_000_000, 29),
    (550_000_000, 30),
    (695_000_000, 31),
    (910_000_000, 32),
    (1_400_000_000, 33),
    (float("inf"), 34),
]

TER_CATEGORY_B = [
    (6_200_000, 0),
    (6_500_000, 0.25),
    (6_850_000, 0.5),
    (7_300_000, 0.75),
    (9_200_000, 1),
    (10_750_000, 1.5),
    (11_250_000, 2),
    (11_600_000, 2.5),
    (12_600_000, 3),
    (13_600_000, 4),
    (14_950_000, 5),
    (16_400_000, 6),
    (18_450_000, 7),
    (21_850_000, 8),
    (26_000_000, 9),
    (27_700_000, 10),
    (29_350_000, 11),
    (31_450_000, 12),
    (33_950_000, 13),
    (37_100_000, 14),
    (41_100_000, 15),
    (45_800_000, 16),
    (49_500_000, 17),
    (53_800_000, 18),
    (58_500_000, 19),
    (64_000_000, 20),
    (71_000_000, 21),
    (80_000_000, 22),
    (93_000_000, 23),
    (109_000_000, 24),
    (129_000_000, 25),
    (163_000_000, 26),
    (211_000_000, 27),
    (374_000_000, 28),
    (459_000_000, 29),
    (555_000_000, 30),
    (704_000_000, 31),
    (957_000_000, 32),
    (1_405_000_000, 33),
    (float("inf"), 34),
]

TER_CATEGORY_C = [
    (6_600_000, 0),
    (6_950_000, 0.25),
    (7_350_000, 0.5),
    (7_800_000, 0.75),
    (8_850_000, 1),
    (9_800_000, 1.25),
    (10_950_000, 1.5),
    (11_200_000, 1.75),
    (12_050_000, 2),
    (12_950_000, 3),
    (14_150_000, 4),
    (15_550_000, 5),
    (17_050_000, 6),
    (19_500_000, 7),
    (22_700_000, 8),
    (26_600_000, 9),
    (28_100_000, 10),
    (30_100_000, 11),
    (32_600_000, 12),
    (35_400_000, 13),
    (38_900_000, 14),
    (43_000_000, 15),
    (47_400_000, 16),
    (51_200_000, 17),
    (55_800_000, 18),
    (60_400_000, 19),
    (66_700_000, 20),
    (74_500_000, 21),
    (83_200_000, 22),
    (95_600_000, 23),
    (110_000_000, 24),
    (134_000_000, 25),
    (169_000_000, 26),
    (221_000_000, 27),
    (390_000_000, 28),
    (463_000_000, 29),
    (561_000_000, 30),
    (709_000_000, 31),
    (965_000_000, 32),
    (1_419_000_000, 33),
    (float("inf"), 34),
]


def _lookup_ter(table: list[tuple[float, float]], income: float) -> float:
    """Look up TER rate from a bracket table."""
    for upper_bound, rate in table:
        if income <= upper_bound:
            return rate
    return table[-1][1]


def _calculate_ter(ptkp_status: str, income: float) -> float:
    """Calculate TER rate based on PTKP category and income."""
    if ptkp_status in PTKP_CATEGORY_A:
        return _lookup_ter(TER_CATEGORY_A, income)
    elif ptkp_status in PTKP_CATEGORY_B:
        return _lookup_ter(TER_CATEGORY_B, income)
    elif ptkp_status in PTKP_CATEGORY_C:
        return _lookup_ter(TER_CATEGORY_C, income)
    else:
        # Default to category A if unknown
        return _lookup_ter(TER_CATEGORY_A, income)


def _calculate_ps17(income: float) -> float:
    """Calculate Pasal 17 progressive rate."""
    if income <= 60_000_000:
        return 5
    elif income <= 250_000_000:
        return 15
    elif income <= 500_000_000:
        return 25
    elif income <= 5_000_000_000:
        return 30
    else:
        return 35


def _calculate_harian(income: float) -> float:
    """Calculate daily worker rate."""
    if income <= 450_000:
        return 0
    elif income <= 2_500_000:
        return 0.5
    else:
        return 0


def _calculate_pesangon(income: float) -> float:
    """Calculate severance pay rate."""
    if income <= 50_000_000:
        return 0
    elif income <= 100_000_000:
        return 5
    elif income <= 500_000_000:
        return 15
    else:
        return 25


def _calculate_pensiun(income: float) -> float:
    """Calculate pension rate."""
    if income <= 50_000_000:
        return 0
    else:
        return 5


# =============================================================================
# Public API
# =============================================================================

def get_deemed(tax_object_code: str) -> int:
    """Get the deemed percentage for a BP21 tax object code."""
    if tax_object_code in BP21_TAX_OBJECT_LOOKUP:
        return BP21_TAX_OBJECT_LOOKUP[tax_object_code][1]
    return 100


def get_bppu_rate(tax_object_code: str) -> float:
    """Get the fixed rate for a BPPU tax object code from the reference table."""
    if tax_object_code in BPPU_TAX_OBJECT_LOOKUP:
        return BPPU_TAX_OBJECT_LOOKUP[tax_object_code][1]
    return 0


def get_tariff_type(tax_object_code: str, doc_type: str = "BP21") -> str:
    """Get the tariff type for a tax object code."""
    lookup = BP21_TAX_OBJECT_LOOKUP if doc_type == "BP21" else BPMP_TAX_OBJECT_LOOKUP
    if tax_object_code in lookup:
        return lookup[tax_object_code][2]
    return "TER"


def calculate_rate(
    tax_object_code: str,
    ptkp_status: str,
    gross: float,
    deemed: float = 100,
    doc_type: str = "BP21",
) -> float:
    """
    Calculate the tax rate for a given set of parameters.

    For BP21: income base = gross * deemed / 100
    For BPMP: income base = gross (no deemed)

    Args:
        tax_object_code: Tax object code (e.g. "21-100-01")
        ptkp_status: PTKP status (e.g. "TK/0", "K/1")
        gross: Gross income
        deemed: Deemed percentage (default 100)
        doc_type: "BP21" or "BPMP"

    Returns:
        Tax rate as a percentage (e.g. 5.0 means 5%)
    """
    tariff_type = get_tariff_type(tax_object_code, doc_type)

    # For BP21, calculate income base using deemed
    if doc_type == "BP21":
        income = gross * deemed / 100
    else:
        # BPMP always uses gross directly (all deemed = 100)
        income = gross

    if tariff_type == "TER":
        return _calculate_ter(ptkp_status, income)
    elif tariff_type == "PS17":
        return _calculate_ps17(income)
    elif tariff_type == "HARIAN":
        return _calculate_harian(income)
    elif tariff_type == "PESANGON":
        return _calculate_pesangon(income)
    elif tariff_type == "PENSIUN":
        return _calculate_pensiun(income)
    else:
        # Fixed rate (e.g. "0", "5", "15")
        try:
            return float(tariff_type)
        except ValueError:
            return 0
