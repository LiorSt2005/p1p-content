"""Tests for check_clock_sync: skew detection and ClockSkewError content."""

from datetime import datetime, timedelta, timezone

import pytest

from trading_bot.execution.api_client import ClockSkewError, check_clock_sync


def _server_fn(offset_ms: float):
    """Return a server_time callable whose result is offset_ms ahead of local UTC."""

    def fn() -> datetime:
        return datetime.now(timezone.utc) + timedelta(milliseconds=offset_ms)

    return fn


def test_clock_sync_passes_when_delta_within_limit():
    # 500ms skew — well below the 1000ms threshold
    check_clock_sync(_server_fn(500))


def test_clock_sync_passes_at_exact_limit():
    # Exactly 1000ms is allowed (delta <= limit)
    check_clock_sync(_server_fn(999))


def test_clock_sync_raises_when_delta_exceeds_limit():
    with pytest.raises(ClockSkewError):
        check_clock_sync(_server_fn(1500))


def test_clock_sync_raises_for_large_negative_skew():
    # Server behind local clock by 2 seconds
    with pytest.raises(ClockSkewError):
        check_clock_sync(_server_fn(-2000))


def test_clock_skew_error_message_includes_w32tm_resync():
    with pytest.raises(ClockSkewError) as exc_info:
        check_clock_sync(_server_fn(2000))
    assert "w32tm /resync" in str(exc_info.value)


def test_clock_skew_error_message_includes_delta():
    with pytest.raises(ClockSkewError) as exc_info:
        check_clock_sync(_server_fn(3000))
    assert "ms" in str(exc_info.value)


def test_clock_sync_handles_naive_server_datetime():
    # Server returns a naive datetime — should be treated as UTC without raising
    def naive_server() -> datetime:
        return datetime.now(timezone.utc).replace(tzinfo=None)  # strip tzinfo

    check_clock_sync(naive_server)  # 0ms skew, must pass
