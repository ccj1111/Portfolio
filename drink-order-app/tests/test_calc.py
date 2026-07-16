import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from calc import OrderLine, compute_settlement


def test_zero_orders_no_division_by_zero():
    result = compute_settlement([], actual_total=0)
    assert result.grand_total_estimate == 0
    assert result.ratio is None
    assert result.total_adjusted == 0
    assert result.buffer == 0


def test_discount_never_costs_organizer_money():
    # grand total 100, actual bill only 90 (discount applied)
    lines = [
        OrderLine(id=1, person_name="Alice", unit_price=60, qty=1),
        OrderLine(id=2, person_name="Bob", unit_price=40, qty=1),
    ]
    result = compute_settlement(lines, actual_total=90)
    assert result.per_order[1] == 54  # ceil(60 * 0.9)
    assert result.per_order[2] == 36  # ceil(40 * 0.9)
    assert result.total_adjusted == 90
    assert result.buffer == 0
    assert result.total_adjusted >= result.actual_total


def test_platform_fee_never_costs_organizer_money():
    # grand total 100, actual bill 110 (foodpanda platform fee)
    lines = [
        OrderLine(id=1, person_name="Alice", unit_price=60, qty=1),
        OrderLine(id=2, person_name="Bob", unit_price=40, qty=1),
    ]
    result = compute_settlement(lines, actual_total=110)
    assert result.per_order[1] == 66  # ceil(60 * 1.1)
    assert result.per_order[2] == 44  # ceil(40 * 1.1)
    assert result.total_adjusted == 110
    assert result.buffer == 0


def test_rounding_always_buffers_in_organizers_favor():
    # 3 lines of 33/33/34 = 100, actual 95 -> ratio 0.95, none divide evenly
    lines = [
        OrderLine(id=1, person_name="Alice", unit_price=33, qty=1),
        OrderLine(id=2, person_name="Bob", unit_price=33, qty=1),
        OrderLine(id=3, person_name="Carol", unit_price=34, qty=1),
    ]
    result = compute_settlement(lines, actual_total=95)
    assert result.per_order[1] == 32  # ceil(31.35)
    assert result.per_order[2] == 32  # ceil(31.35)
    assert result.per_order[3] == 33  # ceil(32.3)
    assert result.total_adjusted == 97
    assert result.buffer == 2
    # Never below actual_total, buffer always >= 0.
    assert result.total_adjusted >= result.actual_total
    assert result.buffer >= 0


def test_multiple_lines_same_person_are_summed():
    lines = [
        OrderLine(id=1, person_name="Alice", unit_price=50, qty=2),  # line total 100
        OrderLine(id=2, person_name="Alice", unit_price=30, qty=1),  # line total 30
    ]
    result = compute_settlement(lines, actual_total=130)
    assert result.per_person["Alice"] == 130


def test_no_actual_total_yet_is_handled_by_caller():
    # Callers should not invoke compute_settlement until actual_total is known;
    # this test just documents that ratio requires an int actual_total.
    lines = [OrderLine(id=1, person_name="Alice", unit_price=50, qty=1)]
    result = compute_settlement(lines, actual_total=50)
    assert result.ratio == 1
    assert result.per_order[1] == 50
    assert result.buffer == 0


def test_exact_fraction_avoids_float_error():
    # A case that would be borderline with float rounding: 3 * (1/3) situations.
    lines = [OrderLine(id=i, person_name=f"P{i}", unit_price=1, qty=1) for i in range(1, 4)]
    result = compute_settlement(lines, actual_total=1)
    # grand total 3, actual 1 -> ratio 1/3, each line ceil(1/3) = 1
    assert result.per_order == {1: 1, 2: 1, 3: 1}
    assert result.total_adjusted == 3
    assert result.buffer == 2
