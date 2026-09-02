from decimal import Decimal

import pytest

from excel_accountant.decimal_codec import (
    AmountParseError,
    AmountRangeError,
    INT64_MAX,
    encode_problem,
    format_decimal,
    parse_amount,
    parse_storage_decimal,
)


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("0", Decimal("0")),
        ("1.2300", Decimal("1.2300")),
        ("-12.34567", Decimal("-12.34567")),
        ("+1,234,567.8901", Decimal("1234567.8901")),
        (" 42.00 ", Decimal("42.00")),
    ],
)
def test_parse_amount_is_exact(text: str, expected: Decimal) -> None:
    result = parse_amount(text)
    assert result == expected
    assert result.as_tuple().exponent == expected.as_tuple().exponent


@pytest.mark.parametrize(
    "text",
    ["", "1,23", "￥12.00", "12元", "NaN", "Infinity", "1e3", "--1"],
)
def test_parse_amount_rejects_ambiguous_text(text: str) -> None:
    with pytest.raises(AmountParseError):
        parse_amount(text)


def test_storage_parser_accepts_scientific_notation_exactly() -> None:
    assert parse_storage_decimal("1.25E-3") == Decimal("0.00125")


def test_encode_problem_only_pads_zeros() -> None:
    problem = encode_problem(
        [Decimal("12.3"), Decimal("0.0456"), Decimal("-2.000")],
        [Decimal("10.3456")],
    )
    assert problem.scale == 10_000
    assert problem.divisor == 8
    assert problem.amounts == (15_375, 57, -2_500)
    assert problem.targets == (12_932,)
    assert problem.restore(problem.targets[0]) == Decimal("10.3456")


def test_encode_problem_reduces_by_gcd_reversibly() -> None:
    problem = encode_problem(
        [Decimal("1.20"), Decimal("2.40")],
        [Decimal("3.60")],
    )
    assert problem.scale == 100
    assert problem.divisor == 120
    assert problem.amounts == (1, 2)
    assert problem.targets == (3,)
    assert problem.restore(3) == Decimal("3.6")


def test_encode_problem_rejects_empty_targets() -> None:
    with pytest.raises(AmountRangeError):
        encode_problem([Decimal("1")], [])


def test_encode_problem_rejects_solver_overflow() -> None:
    with pytest.raises(AmountRangeError):
        encode_problem([Decimal(INT64_MAX), Decimal(1)], [Decimal(1)])


def test_format_decimal_never_uses_exponent() -> None:
    assert format_decimal(Decimal("1E-7")) == "0.0000001"
