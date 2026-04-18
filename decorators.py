import sys
import time
import random
from functools import wraps
from constants import _sp_semaphore, _RETRY_STATUS


def _with_backpressure(fn):
    """Wrap a sync tool: bound concurrency + retry transient throttling."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        with _sp_semaphore:
            max_attempts = 4
            base = 0.5
            for attempt in range(max_attempts):
                try:
                    return fn(*args, **kwargs)
                except Exception as e:
                    msg = str(e)
                    transient = any(s in msg for s in _RETRY_STATUS)
                    if not transient or attempt == max_attempts - 1:
                        raise
                    sleep_s = base * (2 ** attempt) + random.uniform(0, 0.25)
                    print(
                        f"[retry] {fn.__name__} attempt {attempt + 1} after {sleep_s:.2f}s ({msg[:80]})",
                        file=sys.stderr,
                    )
                    time.sleep(sleep_s)
    return wrapper
