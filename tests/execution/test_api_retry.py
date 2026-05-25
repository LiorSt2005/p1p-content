"""Tests for call_with_retry: 429 / 5xx retry logic, backoff, silent-failure guard."""

import ast
import pathlib
from unittest.mock import MagicMock, call, patch

import pytest

from trading_bot.execution.api_client import (
    ApiError,
    RateLimitExhaustedError,
    call_with_retry,
)


def test_successful_call_returns_result_without_retry():
    fn = MagicMock(return_value="ok")
    result = call_with_retry(fn)
    assert result == "ok"
    fn.assert_called_once()


def test_429_retries_using_retry_after():
    retry_after = 7.5
    fn = MagicMock(
        side_effect=[
            ApiError(429, retry_after=retry_after),
            "ok",
        ]
    )
    with patch("trading_bot.execution.api_client.time.sleep") as mock_sleep:
        result = call_with_retry(fn)

    assert result == "ok"
    assert fn.call_count == 2
    mock_sleep.assert_called_once_with(retry_after)


def test_429_falls_back_to_schedule_when_no_retry_after():
    fn = MagicMock(
        side_effect=[
            ApiError(429),  # no retry_after → use schedule[0] = 1s
            "ok",
        ]
    )
    with patch("trading_bot.execution.api_client.time.sleep") as mock_sleep:
        result = call_with_retry(fn)

    assert result == "ok"
    mock_sleep.assert_called_once_with(1)


def test_429_max_failures_raises_api_rate_limit_exhausted():
    fn = MagicMock(side_effect=ApiError(429))
    with patch("trading_bot.execution.api_client.time.sleep"):
        with pytest.raises(RateLimitExhaustedError) as exc_info:
            call_with_retry(fn)

    assert "API_RATE_LIMIT_EXHAUSTED" in str(exc_info.value)
    assert exc_info.value.error_code == "API_RATE_LIMIT_EXHAUSTED"
    assert fn.call_count == 4  # 1 initial + 3 retries


def test_5xx_exponential_backoff_sequence():
    fn = MagicMock(side_effect=ApiError(500))
    with patch("trading_bot.execution.api_client.time.sleep") as mock_sleep:
        with pytest.raises(ApiError) as exc_info:
            call_with_retry(fn)

    assert exc_info.value.status_code == 500
    assert fn.call_count == 4  # 1 initial + 3 retries
    assert mock_sleep.call_args_list == [call(1), call(2), call(4)]


def test_5xx_503_also_retries():
    fn = MagicMock(side_effect=[ApiError(503), ApiError(503), "ok"])
    with patch("trading_bot.execution.api_client.time.sleep"):
        result = call_with_retry(fn)
    assert result == "ok"


def test_no_silent_failure_non_retryable_4xx():
    fn = MagicMock(side_effect=ApiError(400, "Bad Request"))
    with pytest.raises(ApiError) as exc_info:
        call_with_retry(fn)
    assert exc_info.value.status_code == 400
    fn.assert_called_once()  # no retry attempted


def test_no_silent_failure_non_api_exception():
    fn = MagicMock(side_effect=ValueError("unexpected"))
    with pytest.raises(ValueError, match="unexpected"):
        call_with_retry(fn)
    fn.assert_called_once()


def test_no_alpaca_trading_imports_in_api_client():
    root = pathlib.Path(__file__).parents[2]
    src_file = root / "src" / "trading_bot" / "execution" / "api_client.py"
    tree = ast.parse(src_file.read_text(encoding="utf-8"))

    forbidden = {"alpaca", "alpaca_trade_api", "alpaca.trading", "TradingClient", "submit_order"}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                for bad in forbidden:
                    assert bad not in alias.name, f"Forbidden import: {alias.name}"
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            for bad in forbidden:
                assert bad not in module, f"Forbidden import from: {module}"
            for alias in node.names:
                assert alias.name not in forbidden, f"Forbidden name imported: {alias.name}"
