"""Minimal Supabase connectivity check (no third-party deps).

Examples (PowerShell):
    python supabase_connection_check.py --url "https://<project-ref>.supabase.co" --key "<anon-or-service-role-key>"

Or via environment variables:
    $env:SUPABASE_URL = "https://<project-ref>.supabase.co"
    $env:SUPABASE_KEY = "<anon-or-service-role-key>"
    python supabase_connection_check.py

Checks performed:
- Network/TLS reachability via Auth health endpoint (`/auth/v1/health`). Note: Supabase may return 401 if no `apikey` header is provided.
- (Optional) Key validity against PostgREST (`/rest/v1/`) when a key is provided
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class CheckResult:
    ok: bool
    url: str
    auth_health_status: Optional[int] = None
    rest_status: Optional[int] = None
    detail: str = ""


def _normalize_supabase_url(url: str) -> str:
    url = (url or "").strip().rstrip("/")
    if not url:
        return url
    if not (url.startswith("https://") or url.startswith("http://")):
        url = "https://" + url
    return url


def _http_get(url: str, headers: dict[str, str] | None = None, timeout: float = 10.0) -> tuple[int, str]:
    req = urllib.request.Request(url, method="GET")
    if headers:
        for k, v in headers.items():
            req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(4096)  # small peek
            try:
                body = raw.decode("utf-8", errors="replace")
            except Exception:
                body = ""
            return resp.status, body
    except urllib.error.HTTPError as e:
        body = ""
        try:
            raw = e.read(4096)
            body = raw.decode("utf-8", errors="replace")
        except Exception:
            pass
        return e.code, body


def check_supabase_connection(
    url: Optional[str] = None,
    key: Optional[str] = None,
    *,
    timeout: float = 10.0,
) -> CheckResult:
    """Returns a structured result; does not raise on common failures."""

    url = _normalize_supabase_url(url or os.getenv("SUPABASE_URL", ""))

    key = (key or os.getenv("SUPABASE_KEY") or os.getenv("SUPABASE_ANON_KEY") or os.getenv("SUPABASE_SERVICE_ROLE_KEY") or "").strip()

    if not url:
        return CheckResult(
            ok=False,
            url="",
            detail="Missing SUPABASE_URL. Set env var SUPABASE_URL to https://<project-ref>.supabase.co",
        )

    # 1) Basic reachability: Auth health endpoint
    auth_health_url = f"{url}/auth/v1/health"
    try:
        auth_headers = None
        if key:
            auth_headers = {
                "apikey": key,
                "authorization": f"Bearer {key}",
                "accept": "application/json",
            }
        auth_status, auth_body = _http_get(auth_health_url, headers=auth_headers, timeout=timeout)
    except Exception as e:
        return CheckResult(
            ok=False,
            url=url,
            detail=f"Network/TLS failure hitting {auth_health_url}: {e}",
        )

    # Consider anything not-2xx as a failure for reachability.
    if not (200 <= auth_status < 300):
        if not key and auth_status in (401, 403):
            return CheckResult(
                ok=False,
                url=url,
                auth_health_status=auth_status,
                detail=(
                    "Supabase URL reachable, but the Auth health endpoint requires an API key. "
                    "Provide --key or set SUPABASE_KEY to validate access."
                ),
            )
        return CheckResult(
            ok=False,
            url=url,
            auth_health_status=auth_status,
            detail=f"Auth health endpoint returned HTTP {auth_status}. Body: {auth_body[:300]}",
        )

    # 2) Optional: validate key against PostgREST
    rest_status = None
    if key:
        rest_url = f"{url}/rest/v1/"
        headers = {
            "apikey": key,
            "authorization": f"Bearer {key}",
            "accept": "application/json",
        }
        rest_status, rest_body = _http_get(rest_url, headers=headers, timeout=timeout)

        # 401/403 strongly indicates an invalid/insufficient key.
        if rest_status in (401, 403):
            return CheckResult(
                ok=False,
                url=url,
                auth_health_status=auth_status,
                rest_status=rest_status,
                detail=f"Supabase reachable, but key rejected by PostgREST (HTTP {rest_status}). Body: {rest_body[:300]}",
            )

    return CheckResult(
        ok=True,
        url=url,
        auth_health_status=auth_status,
        rest_status=rest_status,
        detail=(
            "Supabase reachable; auth health OK"
            + (f"; PostgREST responded HTTP {rest_status}" if rest_status is not None else "")
        ),
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Check Supabase reachability and (optionally) key validity.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python supabase_connection_check.py --url https://<project-ref>.supabase.co --key <anon-key>\n"
            "  python supabase_connection_check.py --url https://<project-ref>.supabase.co\n"
            "\n"
            "Environment variables (optional): SUPABASE_URL, SUPABASE_KEY\n"
        ),
    )
    parser.add_argument(
        "--url",
        dest="url",
        default=None,
        help=(
            "Supabase project URL (e.g. https://<project-ref>.supabase.co). "
            "If omitted, uses SUPABASE_URL."
        ),
    )
    parser.add_argument(
        "--key",
        dest="key",
        default=None,
        help=(
            "Supabase anon/service key. If omitted, uses "
            "SUPABASE_KEY/SUPABASE_ANON_KEY/SUPABASE_SERVICE_ROLE_KEY."
        ),
    )
    parser.add_argument(
        "--timeout",
        dest="timeout",
        type=float,
        default=10.0,
        help="HTTP timeout in seconds (default: 10).",
    )
    return parser


def main(argv: list[str]) -> int:
    args = _build_parser().parse_args(argv)

    result = check_supabase_connection(url=args.url, key=args.key, timeout=args.timeout)

    payload = {
        "ok": result.ok,
        "url": result.url,
        "auth_health_status": result.auth_health_status,
        "rest_status": result.rest_status,
        "detail": result.detail,
    }
    print(json.dumps(payload, indent=2))
    return 0 if result.ok else 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
