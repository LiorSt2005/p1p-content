"""Generic API retry/backoff wrapper and clock synchronisation check."""

import logging
import time
from collections.abc import Callable
from datetime import datetime, timezone
from typing import Any

logger = logging.getLogger(__name__)

MAX_RETRIES = 3
_BACKOFF_SCHEDULE = [1, 2, 4]  # seconds between successive retry attempts
_CLOCK_SKEW_LIMIT_MS = 1000


class ApiError(Exception):
    """Wraps an HTTP error response from any broker API."""

    def __init__(
        self,
        status_code: int,
        message: str = "",
        retry_after: float | None = None,
    ) -> None:
        self.status_code = status_code
        self.retry_after = retry_after
        super().__init__(message or f"HTTP {status_code}")


class RateLimitExhaustedError(Exception):
    """All retry attempts on HTTP 429 have been consumed."""

    error_code = "API_RATE_LIMIT_EXHAUSTED"


class ClockSkewError(Exception):
    """Local clock deviates from server time beyond the acceptable threshold."""


def call_with_retry(
    fn: Callable[[], Any],
    *,
    max_retries: int = MAX_RETRIES,
) -> Any:
    """
    Generic retry/backoff wrapper for API calls.

    Behaviour:
    - HTTP 429: reads retry_after from ApiError; falls back to _BACKOFF_SCHEDULE.
      After max_retries exhausted raises RateLimitExhaustedError (error_code
      API_RATE_LIMIT_EXHAUSTED) and logs the failure.
    - HTTP 5xx: exponential backoff using _BACKOFF_SCHEDULE [1s, 2s, 4s].
      After max_retries exhausted re-raises the original ApiError and logs.
    - Any other exception: re-raised immediately — no silent failures.
    - A successful call returns the result without any retry overhead.
    """
    for attempt in range(max_retries + 1):
        try:
            return fn()
        except ApiError as exc:
            is_last_attempt = attempt >= max_retries

            if exc.status_code == 429:
                if is_last_attempt:
                    logger.error(
                        "HTTP 429: API_RATE_LIMIT_EXHAUSTED after %d retries.",
                        max_retries,
                    )
                    raise RateLimitExhaustedError(
                        f"API_RATE_LIMIT_EXHAUSTED after {max_retries} retries"
                    ) from exc
                wait = (
                    exc.retry_after
                    if exc.retry_after is not None
                    else _BACKOFF_SCHEDULE[attempt]
                )
                logger.warning(
                    "HTTP 429 rate limited (attempt %d/%d). Waiting %.1fs before retry.",
                    attempt + 1,
                    max_retries + 1,
                    wait,
                )
                time.sleep(wait)

            elif 500 <= exc.status_code < 600:
                if is_last_attempt:
                    logger.error(
                        "HTTP %d: server error exhausted after %d attempts.",
                        exc.status_code,
                        max_retries + 1,
                    )
                    raise
                wait = _BACKOFF_SCHEDULE[attempt]
                logger.warning(
                    "HTTP %d server error (attempt %d/%d). Backing off %.1fs.",
                    exc.status_code,
                    attempt + 1,
                    max_retries + 1,
                    wait,
                )
                time.sleep(wait)

            else:
                raise

    raise RuntimeError("Unreachable")  # pragma: no cover


def check_clock_sync(server_time_fn: Callable[[], datetime]) -> None:
    """
    Compare local UTC time against broker/server time.

    delta <= 1000 ms  →  passes silently.
    delta >  1000 ms  →  raises ClockSkewError with Windows remediation hint.

    The bot must NOT retry automatically after this error; startup aborts
    until the operator runs: w32tm /resync
    """
    server_time = server_time_fn()
    local_time = datetime.now(timezone.utc)

    if server_time.tzinfo is None:
        server_time = server_time.replace(tzinfo=timezone.utc)

    delta_ms = abs((local_time - server_time).total_seconds() * 1000)

    if delta_ms > _CLOCK_SKEW_LIMIT_MS:
        raise ClockSkewError(
            f"Clock skew {delta_ms:.0f}ms exceeds {_CLOCK_SKEW_LIMIT_MS}ms limit. "
            f"Fix with: w32tm /resync"
        )

    logger.debug("Clock sync OK: delta %.0fms.", delta_ms)
