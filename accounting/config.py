"""Settings. Every value comes from the environment (the VPS `.env`); nothing here is a secret.

Naming follows the usage repo so the same `.env` conventions carry over
(SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASSWORD / EMAIL_FROM).
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(os.environ.get("ACCOUNTING_ENV_FILE", ".env"))


def _bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")


def _int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    try:
        return int(raw) if raw not in (None, "") else default
    except ValueError:
        return default


def _list(name: str, default: str = "") -> list[str]:
    raw = os.environ.get(name, default)
    return [p.strip() for p in raw.split(",") if p.strip()]


@dataclass(frozen=True)
class Settings:
    # --- hosting ---------------------------------------------------------
    portal_host: str = os.environ.get("PORTAL_HOST", "accounting.90ten.life")
    portal_base_url: str = os.environ.get("PORTAL_BASE_URL", "https://accounting.90ten.life")
    secure_cookies: bool = _bool("SECURE_COOKIES", True)
    environment: str = os.environ.get("ACCOUNTING_ENV", "production")  # production | test

    # --- portal database (Supabase project "maxx-accounting" or local Postgres) ---
    database_url: str = os.environ.get("ACCOUNTING_DATABASE_URL", "postgresql://ap_test:ap_test@localhost/ap_test")
    db_pool_min: int = _int("ACCOUNTING_DB_POOL_MIN", 1)
    db_pool_max: int = _int("ACCOUNTING_DB_POOL_MAX", 8)
    # Column encryption key for TINs and bank fingerprints (32 bytes, base64). Held by the app, not the DB.
    field_key_b64: str = os.environ.get("ACCOUNTING_FIELD_KEY", "")

    # --- object store for originals and renders -----------------------------
    # Local directory in testing; Supabase Storage (S3-compatible) in production.
    store_backend: str = os.environ.get("ACCOUNTING_STORE_BACKEND", "local")  # local | s3
    store_local_dir: Path = Path(os.environ.get("ACCOUNTING_STORE_DIR", "var/store"))
    store_s3_endpoint: str = os.environ.get("ACCOUNTING_STORE_S3_ENDPOINT", "")
    store_s3_bucket_originals: str = os.environ.get("ACCOUNTING_STORE_BUCKET_ORIGINALS", "originals")
    store_s3_bucket_renders: str = os.environ.get("ACCOUNTING_STORE_BUCKET_RENDERS", "renders")
    store_s3_access_key: str = os.environ.get("ACCOUNTING_STORE_S3_ACCESS_KEY", "")
    store_s3_secret_key: str = os.environ.get("ACCOUNTING_STORE_S3_SECRET_KEY", "")

    # --- identity: the MaxxDash Supabase project (identity ONLY) --------------
    maxxdash_supabase_url: str = os.environ.get("MAXXDASH_SUPABASE_URL", "")
    maxxdash_anon_key: str = os.environ.get("MAXXDASH_SUPABASE_ANON_KEY", "")
    maxxdash_dashboard_url: str = os.environ.get("MAXXDASH_DASHBOARD_URL", "https://dashboard.90ten.life")
    jwks_cache_seconds: int = _int("JWKS_CACHE_SECONDS", 3600)
    session_idle_minutes: int = _int("SESSION_IDLE_MINUTES", 30)
    session_absolute_hours: int = _int("SESSION_ABSOLUTE_HOURS", 8)
    session_revalidate_minutes: int = _int("SESSION_REVALIDATE_MINUTES", 60)
    signin_rate_limit_per_15min: int = _int("SIGNIN_RATE_LIMIT", 10)

    # --- sibling apps shown as tiles ----------------------------------------
    usage_url: str = os.environ.get("USAGE_URL", "https://usage.accounting.90ten.life")
    commissions_generator_url: str = os.environ.get("COMMISSIONS_GENERATOR_URL", "https://commissions.accounting.90ten.life")

    # --- mail: accounts@90ten.life on Hostinger ------------------------------
    imap_host: str = os.environ.get("IMAP_HOST", "imap.hostinger.com")
    imap_port: int = _int("IMAP_PORT", 993)
    imap_user: str = os.environ.get("IMAP_USER", "accounts@90ten.life")
    imap_password: str = os.environ.get("IMAP_PASSWORD", "")
    imap_folder: str = os.environ.get("IMAP_FOLDER", "INBOX")
    imap_poll_seconds: int = _int("IMAP_POLL_SECONDS", 60)
    smtp_host: str = os.environ.get("SMTP_HOST", "smtp.hostinger.com")
    smtp_port: int = _int("SMTP_PORT", 465)
    smtp_user: str = os.environ.get("SMTP_USER", "accounts@90ten.life")
    smtp_password: str = os.environ.get("SMTP_PASSWORD", "")
    email_from: str = os.environ.get("EMAIL_FROM", "accounts@90ten.life")
    # INT-4: the only path we trust. Microsoft 365 outbound hosts and the tenant that must appear in
    # Authentication-Results. Anything else arriving at accounts@ is quarantined as a direct sender.
    forwarder_host_suffixes: list[str] = field(default_factory=lambda: _list(
        "FORWARDER_HOST_SUFFIXES", ".outbound.protection.outlook.com,.protection.outlook.com"))
    forwarder_authserv_id: str = os.environ.get("FORWARDER_AUTHSERV_ID", "maxxortho.com")
    heartbeat_min_per_day: int = _int("HEARTBEAT_MIN_PER_DAY", 1)
    heartbeat_max_per_day: int = _int("HEARTBEAT_MAX_PER_DAY", 200)

    # --- people ----------------------------------------------------------------
    controller_email: str = os.environ.get("CONTROLLER_EMAIL", "")
    backup_controller_email: str = os.environ.get("BACKUP_CONTROLLER_EMAIL", "corey.perine@maxxortho.com")
    system_owner_email: str = os.environ.get("SYSTEM_OWNER_EMAIL", "lenny.knobbs@maxxortho.com")
    cfo_email: str = os.environ.get("CFO_EMAIL", "")
    security_email: str = os.environ.get("SECURITY_EMAIL", "")
    timezone: str = os.environ.get("ACCOUNTING_TIMEZONE", "America/New_York")

    # --- parser container ----------------------------------------------------
    parser_socket: str = os.environ.get("PARSER_SOCKET", "var/parser.sock")
    parser_max_pages: int = _int("PARSER_MAX_PAGES", 50)
    parser_max_bytes: int = _int("PARSER_MAX_BYTES", 100 * 1024 * 1024)
    parser_cpu_seconds: int = _int("PARSER_CPU_SECONDS", 60)

    # --- capture (Anthropic only — no OpenRouter for invoices, see spec §6) --------
    anthropic_api_key: str = os.environ.get("ANTHROPIC_API_KEY", "")
    anthropic_model: str = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
    anthropic_escalation_model: str = os.environ.get("ANTHROPIC_ESCALATION_MODEL", "claude-opus-5-5")
    capture_confidence_threshold: float = float(os.environ.get("CAPTURE_CONFIDENCE_THRESHOLD", "0.85"))

    # --- approvals -----------------------------------------------------------
    rp_id: str = os.environ.get("WEBAUTHN_RP_ID", "accounting.90ten.life")
    rp_name: str = os.environ.get("WEBAUTHN_RP_NAME", "Maxx Accounting Portal")
    signing_session_idle_minutes: int = _int("SIGNING_SESSION_IDLE_MINUTES", 15)
    signing_session_max_minutes: int = _int("SIGNING_SESSION_MAX_MINUTES", 60)
    # Signing service in Lenny's AWS account (Lambda function URL, IAM auth). Empty = local dev signer.
    signing_service_url: str = os.environ.get("SIGNING_SERVICE_URL", "")
    aws_region: str = os.environ.get("AWS_REGION", "us-east-1")
    aws_access_key_id: str = os.environ.get("AWS_ACCESS_KEY_ID", "")
    aws_secret_access_key: str = os.environ.get("AWS_SECRET_ACCESS_KEY", "")
    # Dev-only seal key (PEM, ECDSA P-256). Never set in production; production fails closed without the service.
    dev_seal_private_key_pem: str = os.environ.get("DEV_SEAL_PRIVATE_KEY_PEM", "")

    # --- audit anchoring -----------------------------------------------------
    anchor_bucket: str = os.environ.get("ANCHOR_BUCKET", "")
    anchor_interval_minutes: int = _int("ANCHOR_INTERVAL_MINUTES", 15)

    # --- business rules (defaults; live values are rows in `config`) --------------
    default_step_up_amount_cents: int = _int("STEP_UP_AMOUNT_CENTS", 2_500_000)
    recurring_tolerance_pct: float = float(os.environ.get("RECURRING_TOLERANCE_PCT", "2.0"))
    second_person_release_cents: int = _int("SECOND_PERSON_RELEASE_CENTS", 1_000_000)

    @property
    def is_test(self) -> bool:
        return self.environment == "test"


settings = Settings()
