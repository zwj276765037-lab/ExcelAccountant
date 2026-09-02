from __future__ import annotations

import math
import re
from decimal import Decimal, InvalidOperation
from functools import reduce
from typing import Iterable, Sequence

from .models import ScaledProblem

INT64_MAX = (1 << 63) - 1

_PLAIN_AMOUNT_RE = re.compile(
    r"^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$"
)
_STORAGE_AMOUNT_RE = re.compile(
    r"^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?$"
)


class AmountParseError(ValueError):
    """Raised when a string cannot be interpreted as an exact amount."""


class AmountRangeError(ValueError):
    """Raised when an exact scaled problem cannot fit the solver range."""


def parse_amount(text: str) -> Decimal:
    """Parse a user or strict text-cell amount without float conversion."""

    value = text.strip()
    if not value or not _PLAIN_AMOUNT_RE.fullmatch(value):
        raise AmountParseError(f"无效金额格式：{text!r}")
    return _to_finite_decimal(value.replace(",", ""), text)


def parse_storage_decimal(text: str) -> Decimal:
    """Parse the exact numeric token stored in XLSX XML."""

    value = text.strip()
    if not value or not _STORAGE_AMOUNT_RE.fullmatch(value):
        raise AmountParseError(f"无效工作簿数值：{text!r}")
    return _to_finite_decimal(value, text)


def _to_finite_decimal(value: str, original: str) -> Decimal:
    try:
        result = Decimal(value)
    except InvalidOperation as exc:
        raise AmountParseError(f"无效金额：{original!r}") from exc
    if not result.is_finite():
        raise AmountParseError(f"金额必须是有限十进制数：{original!r}")
    return result


def decimal_places(value: Decimal) -> int:
    exponent = value.as_tuple().exponent
    return max(0, -exponent)


def scale_decimal(value: Decimal, scale: int) -> int:
    scaled = value * scale
    integral = scaled.to_integral_value()
    if scaled != integral:
        raise AmountRangeError(
            f"金额 {value} 无法在尺度 {scale} 下无损整数化"
        )
    return int(integral)


def encode_problem(
    amounts: Sequence[Decimal], targets: Sequence[Decimal]
) -> ScaledProblem:
    if not targets:
        raise AmountRangeError("至少需要一个目标金额")
    all_values = tuple(amounts) + tuple(targets)
    max_places = max((decimal_places(value) for value in all_values), default=0)
    scale = 10**max_places
    raw_amounts = tuple(scale_decimal(value, scale) for value in amounts)
    raw_targets = tuple(scale_decimal(value, scale) for value in targets)
    divisor = _greatest_common_divisor(raw_amounts + raw_targets)
    scaled_amounts = tuple(value // divisor for value in raw_amounts)
    scaled_targets = tuple(value // divisor for value in raw_targets)
    _validate_solver_range(scaled_amounts, scaled_targets)
    return ScaledProblem(
        scale=scale,
        divisor=divisor,
        amounts=scaled_amounts,
        targets=scaled_targets,
    )


def _greatest_common_divisor(values: Iterable[int]) -> int:
    nonzero = (abs(value) for value in values if value)
    return reduce(math.gcd, nonzero, 0) or 1


def _validate_solver_range(amounts: Sequence[int], targets: Sequence[int]) -> None:
    for value in (*amounts, *targets):
        if abs(value) > INT64_MAX:
            raise AmountRangeError("金额超出精确求解器的 64 位整数范围")
    absolute_sum = sum(abs(value) for value in amounts)
    if absolute_sum > INT64_MAX:
        raise AmountRangeError("候选金额绝对值总和超出精确求解器范围")


def format_decimal(value: Decimal) -> str:
    """Render without exponent notation and without changing precision."""

    return format(value, "f")
