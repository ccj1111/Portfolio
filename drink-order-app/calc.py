"""Settlement / cost-sharing math.

Implements spec section 7.3: whatever the actual bill turns out to be
(discount applied or platform fee tacked on), the organizer must never end
up out of pocket. Every order line is scaled by the same ratio and then
rounded UP (never down), so any rounding slack always favors the organizer.

All money amounts are whole NT dollars (ints). The ratio is kept as an
exact Fraction so there is no floating-point error to compensate for.
"""

from dataclasses import dataclass
from fractions import Fraction
from math import ceil


@dataclass(frozen=True)
class OrderLine:
    id: int
    person_name: str
    unit_price: int
    qty: int

    @property
    def line_total(self) -> int:
        return self.unit_price * self.qty


@dataclass(frozen=True)
class SettlementResult:
    grand_total_estimate: int
    actual_total: int
    ratio: Fraction | None
    per_order: dict  # order id -> adjusted line total (int)
    per_person: dict  # person name -> adjusted total (int)
    total_adjusted: int
    buffer: int  # total_adjusted - actual_total, always >= 0


def compute_settlement(lines, actual_total: int) -> SettlementResult:
    """Compute per-line/per-person amounts actually owed.

    lines: iterable of OrderLine
    actual_total: the real amount the organizer paid (int, NT dollars)
    """
    lines = list(lines)
    grand_total_estimate = sum(line.line_total for line in lines)

    per_order: dict = {}
    per_person: dict = {}

    if grand_total_estimate <= 0:
        # No orders to prorate against - avoid division by zero (spec 7.3 note).
        return SettlementResult(
            grand_total_estimate=grand_total_estimate,
            actual_total=actual_total,
            ratio=None,
            per_order=per_order,
            per_person=per_person,
            total_adjusted=0,
            buffer=0,
        )

    ratio = Fraction(actual_total, grand_total_estimate)

    for line in lines:
        adjusted = ceil(Fraction(line.line_total) * ratio)
        per_order[line.id] = adjusted
        per_person[line.person_name] = per_person.get(line.person_name, 0) + adjusted

    total_adjusted = sum(per_order.values())
    buffer = total_adjusted - actual_total

    return SettlementResult(
        grand_total_estimate=grand_total_estimate,
        actual_total=actual_total,
        ratio=ratio,
        per_order=per_order,
        per_person=per_person,
        total_adjusted=total_adjusted,
        buffer=buffer,
    )
