import os
import sys
import json
import time
import hashlib
import inspect
from functools import wraps

_REDIS_URL = os.getenv("REDIS_URL", "")
_CACHE_ENABLED = os.getenv("CACHE_ENABLED", "true").lower() == "true"
_DEFAULT_TTL = int(os.getenv("CACHE_DEFAULT_TTL", "300"))
_MAX_BYTES = int(os.getenv("CACHE_MAX_BYTES", str(1 * 1024 * 1024)))
_NAMESPACE = os.getenv("CACHE_NAMESPACE", "sp:v1")

TTL_MAP: dict[str, int] = {
    "list_files":        300,
    "read_file_content": 1800,
    "search_files":      600,
    "search_content":    600,
    "read_pdf":          86400,
    "read_docx":         86400,
    "read_pptx":         86400,
    "get_file_metadata": 30,
}

# In-process stats counters
_hits = 0
_misses = 0
_bypasses = 0
_bytes_served = 0

# Redis client state
_redis_client = None
_client_healthy = True
_last_probe: float = 0.0
_PROBE_INTERVAL = 30.0


def _get_redis():
    global _redis_client, _client_healthy, _last_probe

    if not _REDIS_URL or not _CACHE_ENABLED:
        return None

    if _redis_client is not None and _client_healthy:
        return _redis_client

    now = time.monotonic()
    if not _client_healthy and (now - _last_probe) < _PROBE_INTERVAL:
        return None

    try:
        import redis
        client = redis.Redis.from_url(
            _REDIS_URL,
            decode_responses=True,
            socket_timeout=0.25,
            socket_connect_timeout=0.5,
            health_check_interval=30,
            retry_on_timeout=False,
        )
        client.ping()
        _redis_client = client
        _client_healthy = True
        return _redis_client
    except Exception as e:
        _client_healthy = False
        _last_probe = now
        print(f"[cache] BYPASS: redis unreachable ({e})", file=sys.stderr)
        return None


def _make_key(fn_name: str, bound_args: inspect.BoundArguments) -> str:
    bound_args.apply_defaults()
    canonical = json.dumps(bound_args.arguments, sort_keys=True, default=str, ensure_ascii=False)
    digest = hashlib.sha1(canonical.encode()).hexdigest()[:16]
    return f"{_NAMESPACE}:{fn_name}:{digest}"


def _is_error(fn_name: str, result: str) -> bool:
    return result.startswith(f"[{fn_name}]") or result.startswith("[binary content")


def _with_cache(fn):
    sig = inspect.signature(fn)

    @wraps(fn)
    def wrapper(*args, **kwargs):
        global _hits, _misses, _bypasses, _bytes_served

        client = _get_redis()
        if client is None:
            _bypasses += 1
            return fn(*args, **kwargs)

        try:
            bound = sig.bind(*args, **kwargs)
            key = _make_key(fn.__name__, bound)
        except Exception:
            _bypasses += 1
            return fn(*args, **kwargs)

        # GET
        try:
            cached = client.get(key)
            if cached is not None:
                _hits += 1
                _bytes_served += len(cached)
                print(f"[cache HIT] {fn.__name__}", file=sys.stderr)
                return cached
        except Exception as e:
            global _client_healthy
            _client_healthy = False
            print(f"[cache] BYPASS: {e}", file=sys.stderr)
            _bypasses += 1
            return fn(*args, **kwargs)

        _misses += 1
        print(f"[cache MISS] {fn.__name__}", file=sys.stderr)
        result = fn(*args, **kwargs)

        # SET — skip errors and oversized results
        if _is_error(fn.__name__, result):
            return result
        if len(result) > _MAX_BYTES:
            return result

        ttl = TTL_MAP.get(fn.__name__, _DEFAULT_TTL)
        try:
            client.set(key, result, ex=ttl)
            _bytes_served += len(result)
        except Exception as e:
            _client_healthy = False
            print(f"[cache] SET failed: {e}", file=sys.stderr)

        return result

    return wrapper


def get_stats() -> dict:
    return {
        "hits": _hits,
        "misses": _misses,
        "bypasses": _bypasses,
        "bytes_served": _bytes_served,
        "redis_healthy": _client_healthy,
        "redis_url": _REDIS_URL or "(not configured)",
        "cache_enabled": _CACHE_ENABLED,
    }
