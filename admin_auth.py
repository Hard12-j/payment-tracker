"""Admin authentication for the ?admin page.

* First use        : "Set your admin password" -> OTP is e-mailed to ADMIN_EMAIL -> the admin
                     chooses their own password. Nobody else ever knows it.
* Login            : password is checked against a salted scrypt hash (never stored in plain text).
* Forgot password  : same OTP flow, started from the login screen.
* Change password  : for a logged-in admin (current password + new password).
* Reset / change   : logs the admin out of ALL other devices (within ~20 seconds).

Where things live
* ADMIN_EMAIL      : in Secrets (the only address an OTP is ever sent to).
* password hash    : in the "Admin Auth" tab of the same Google Sheet / Excel file as the payments,
                     so it survives reboots on Streamlit Cloud. Keep that Sheet private.
* OTP              : only in the requesting browser session (hashed, 10 min, 5 tries, single use).

Backwards compatible: if ADMIN_EMAIL is NOT set, the old ADMIN_PASSWORD secret still works.
"""
import hashlib
import hmac
import secrets
import smtplib
import threading
import time
from email.message import EmailMessage

import streamlit as st

import store
from settings import ORG_NAME, SMTP_HOST, SMTP_PASSWORD, SMTP_PORT, SMTP_USER, get_setting

AUTH_SHEET = "Admin Auth"
AUTH_HEADERS = ["Key", "Value"]

OTP_TTL = 10 * 60            # seconds an OTP stays valid
OTP_MAX_TRIES = 5            # wrong guesses before the OTP is thrown away
RESEND_COOLDOWN = 60         # seconds between OTP e-mails (per browser session)
MAX_OTP_PER_HOUR = 5         # OTP e-mails per hour (whole app)
MAX_LOGIN_FAILS = 5          # wrong passwords in a row ...
LOCKOUT_SECONDS = 5 * 60     # ... then login is paused this long (whole app)
MIN_PASSWORD_LEN = 8
SESSION_CHECK_SECONDS = 20    # how often a logged-in session re-checks it was not revoked

_COMMON = {
    "password", "password1", "password123", "pass@1234", "pass1234", "admin123", "admin@123",
    "admin1234", "welcome123", "qwerty123", "12345678", "123456789", "1234567890", "iloveyou",
}

_state_lock = threading.Lock()
_login_fail = {"count": 0, "locked_until": 0.0}
_otp_sends: list = []
_ready_for = None            # backend object whose "Admin Auth" tab we already checked


# --------------------------------------------------------------------------
# Password hashing (standard library only)
# --------------------------------------------------------------------------
def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.scrypt(password.encode(), salt=salt, n=2 ** 14, r=8, p=1, dklen=32)
    return f"scrypt${salt.hex()}${dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        scheme, salt_hex, hash_hex = stored.split("$")
        if scheme != "scrypt":
            return False
        dk = hashlib.scrypt(password.encode(), salt=bytes.fromhex(salt_hex),
                            n=2 ** 14, r=8, p=1, dklen=32)
        return hmac.compare_digest(dk.hex(), hash_hex)
    except Exception:
        return False


def password_problem(pw: str):
    """Return a human-readable problem with the password, or None if it is acceptable."""
    if len(pw) < MIN_PASSWORD_LEN:
        return f"Password must be at least {MIN_PASSWORD_LEN} characters."
    if not any(c.isalpha() for c in pw) or not any(c.isdigit() for c in pw):
        return "Password must contain at least one letter and one number."
    if pw.strip().lower() in _COMMON:
        return "That password is too common. Please choose a different one."
    return None


# --------------------------------------------------------------------------
# OTP (pure functions; the state dict lives in st.session_state)
# --------------------------------------------------------------------------
def _otp_hash(code: str, salt: str) -> str:
    return hashlib.sha256((salt + code).encode()).hexdigest()


def new_otp():
    """Return (plain 6-digit code, state dict to keep in the session)."""
    code = f"{secrets.randbelow(10 ** 6):06d}"
    salt = secrets.token_hex(8)
    return code, {"salt": salt, "hash": _otp_hash(code, salt),
                  "exp": time.time() + OTP_TTL, "tries": 0}


def check_otp(state: dict, code: str) -> str:
    """'ok' | 'wrong' | 'expired' | 'locked'.  Counts wrong guesses in state['tries']."""
    if time.time() > state["exp"]:
        return "expired"
    if state["tries"] >= OTP_MAX_TRIES:
        return "locked"
    if hmac.compare_digest(_otp_hash(code, state["salt"]), state["hash"]):
        return "ok"
    state["tries"] += 1
    return "locked" if state["tries"] >= OTP_MAX_TRIES else "wrong"


def mask_email(email: str) -> str:
    name, _, domain = email.partition("@")
    return f"{name[:1]}***@{domain}" if domain else "***"


def _send_otp_email(to: str, code: str, purpose: str):
    if not (SMTP_USER and SMTP_PASSWORD):
        raise RuntimeError("E-mail is not configured (SMTP_USER / SMTP_PASSWORD in Secrets).")
    msg = EmailMessage()
    msg["Subject"] = f"{ORG_NAME} - admin OTP"
    msg["From"] = f"{ORG_NAME} <{SMTP_USER}>"
    msg["To"] = to
    msg.set_content(
        f"Your OTP to {purpose} is: {code}\n\n"
        f"It is valid for {OTP_TTL // 60} minutes and can be used once.\n"
        "If you did not ask for this, ignore this e-mail - your password has not changed.\n\n"
        f"{ORG_NAME}")
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)


# --------------------------------------------------------------------------
# Storage of the password hash ("Admin Auth" tab: Key | Value)
# --------------------------------------------------------------------------
def _backend():
    global _ready_for
    be = store.get_backend()
    if _ready_for is not be:
        store.ensure_schema(be)
        be.ensure_sheet(AUTH_SHEET, AUTH_HEADERS)
        _ready_for = be
    return be


def _get(key: str, fresh: bool = False) -> str:
    rows = _backend().read(AUTH_SHEET, fresh=fresh)
    for r in rows[1:]:
        if r and r[0] == key and len(r) > 1:
            return r[1]
    return ""


def _set(key: str, value: str):
    with store._lock:
        be = _backend()
        rows = be.read(AUTH_SHEET, fresh=True)
        for i, r in enumerate(rows[1:], start=2):
            if r and r[0] == key:
                be.update_cells(AUTH_SHEET, [(i, 2, value)])
                return
        be.append(AUTH_SHEET, [key, value])


def _save_password(pw: str) -> str:
    """Save the new password hash and rotate the session token, which signs out every device
    that logged in with the old password. Returns the new token."""
    token = secrets.token_hex(16)
    _set("pw_hash", hash_password(pw))
    _set("pw_updated", store.now_str())
    _set("session_token", token)
    return token


def _session_still_valid() -> bool:
    """A logged-in session carries the token that was current when it logged in. If the stored
    token has changed (password reset/changed elsewhere) the session is ended."""
    now = time.time()
    if now - st.session_state.get("auth_checked", 0) < SESSION_CHECK_SECONDS:
        return True
    try:
        current = _get("session_token", fresh=True)
    except Exception:
        return True              # a temporary Sheets error must not lock the admin out
    if current != st.session_state.get("auth_token", ""):
        st.session_state.clear()
        st.session_state["auth_msg"] = ("You were signed out because the admin password was "
                                        "changed. Please log in again.")
        return False
    st.session_state["auth_checked"] = now
    return True


# --------------------------------------------------------------------------
# Throttling
# --------------------------------------------------------------------------
def _lock_remaining() -> int:
    with _state_lock:
        return max(0, int(_login_fail["locked_until"] - time.time()) + 1) \
            if _login_fail["locked_until"] > time.time() else 0


def _note_failure():
    with _state_lock:
        _login_fail["count"] += 1
        if _login_fail["count"] >= MAX_LOGIN_FAILS:
            _login_fail["count"] = 0
            _login_fail["locked_until"] = time.time() + LOCKOUT_SECONDS


def _note_success():
    with _state_lock:
        _login_fail["count"] = 0
        _login_fail["locked_until"] = 0.0


def _otp_request_problem():
    now = time.time()
    wait = RESEND_COOLDOWN - (now - st.session_state.get("otp_last_sent", 0))
    if wait > 0:
        return f"Please wait {int(wait) + 1} seconds before asking for another OTP."
    with _state_lock:
        _otp_sends[:] = [t for t in _otp_sends if now - t < 3600]
        if len(_otp_sends) >= MAX_OTP_PER_HOUR:
            return "Too many OTP requests. Please try again in an hour."
        _otp_sends.append(now)
    return None


# --------------------------------------------------------------------------
# Screens
# --------------------------------------------------------------------------
def _request_otp(admin_email: str, setup: bool):
    problem = _otp_request_problem()
    if problem:
        st.error(problem)
        return
    code, state = new_otp()
    try:
        _send_otp_email(admin_email, code,
                        "set your admin password" if setup else "reset your admin password")
    except Exception as e:
        st.error(f"Could not send the OTP e-mail: {e}")
        return
    st.session_state["otp"] = state
    st.session_state["otp_last_sent"] = time.time()
    st.rerun()


def _reset_ui(admin_email: str, setup: bool):
    st.title("🔐 Set your admin password" if setup else "🔑 Reset admin password")
    if setup:
        st.info("First time here. Choose your own admin password - only you will know it.")
    masked = mask_email(admin_email)
    otp = st.session_state.get("otp")

    if not otp:
        st.write(f"We will send a 6-digit OTP to **{masked}**.")
        c1, c2, _ = st.columns([1, 1, 3])
        if c1.button("Send OTP", type="primary"):
            _request_otp(admin_email, setup)
        if not setup and c2.button("← Back"):
            st.session_state.pop("auth_view", None)
            st.rerun()
        return

    st.success(f"OTP sent to {masked}. It is valid for {OTP_TTL // 60} minutes.")
    with st.form("reset_form"):
        code = st.text_input("OTP from your e-mail", max_chars=6)
        pw1 = st.text_input("New password", type="password")
        pw2 = st.text_input("Confirm new password", type="password")
        ok = st.form_submit_button("Save password")
    if ok:
        problem = password_problem(pw1)
        if problem:
            st.error(problem)
        elif pw1 != pw2:
            st.error("The two passwords do not match.")
        else:
            status = check_otp(otp, code.strip())
            if status == "ok":
                try:
                    _save_password(pw1)
                except Exception as e:
                    st.error(f"Could not save the password: {e}")
                else:
                    for k in ("otp", "auth_view"):
                        st.session_state.pop(k, None)
                    st.session_state["auth_msg"] = "✅ Password saved. Please log in."
                    _note_success()
                    st.rerun()
            elif status == "wrong":
                st.error(f"Wrong OTP. {OTP_MAX_TRIES - otp['tries']} tries left.")
            else:
                st.session_state.pop("otp", None)
                st.error("That OTP has expired or was tried too often. Please request a new one.")
                if st.button("Request a new OTP"):
                    st.rerun()
    c1, c2, _ = st.columns([1, 1, 3])
    if c1.button("Resend OTP"):
        _request_otp(admin_email, setup)
    if not setup and c2.button("← Back to login"):
        for k in ("otp", "auth_view"):
            st.session_state.pop(k, None)
        st.rerun()


def _login_ui():
    st.title("🔒 Admin login")
    msg = st.session_state.pop("auth_msg", None)
    if msg:
        st.success(msg)
    with st.form("login"):
        typed = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
    if ok:
        left = _lock_remaining()
        if left:
            st.error(f"Too many wrong attempts. Try again in {left // 60 + 1} minute(s).")
        else:
            try:
                stored = _get("pw_hash", fresh=True)
            except Exception as e:
                st.error(f"Could not read the saved password: {e}")
                return
            if stored and verify_password(typed, stored):
                _note_success()
                try:
                    st.session_state["auth_token"] = _get("session_token", fresh=True)
                except Exception:
                    st.session_state["auth_token"] = ""
                st.session_state["auth_checked"] = time.time()
                st.session_state["is_admin"] = True
                st.rerun()
            else:
                _note_failure()
                time.sleep(1)  # slow down password guessing
                st.error("Wrong password.")
    if st.button("Forgot password?"):
        st.session_state["auth_view"] = "forgot"
        st.session_state.pop("otp", None)
        st.rerun()


def _legacy_login(password: str) -> bool:
    """Old behaviour: a fixed ADMIN_PASSWORD in Secrets (used only when ADMIN_EMAIL is not set)."""
    st.title("🔒 Admin login")
    with st.form("login"):
        typed = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
    if ok:
        if hmac.compare_digest(typed.encode(), password.encode()):
            st.session_state["is_admin"] = True
            st.rerun()
        else:
            time.sleep(1)
            st.error("Wrong password.")
    return False


# --------------------------------------------------------------------------
# Public API used by admin_page.py
# --------------------------------------------------------------------------
def require_admin() -> bool:
    """True when the visitor is a logged-in admin; otherwise draws the right screen."""
    admin_email = get_setting("ADMIN_EMAIL").strip()
    if st.session_state.get("is_admin"):
        if not admin_email or _session_still_valid():
            return True
        # session revoked: fall through to the login screen (message already set)
    if not admin_email:
        legacy = get_setting("ADMIN_PASSWORD")
        if not legacy:
            st.error("Admin area is disabled. Set ADMIN_EMAIL in the app Secrets "
                     "(.streamlit/secrets.toml).")
            return False
        return _legacy_login(legacy)
    try:
        stored = _get("pw_hash")
    except Exception as e:
        st.error(f"Could not reach the data store: {e}")
        return False
    if not stored:
        _reset_ui(admin_email, setup=True)
    elif st.session_state.get("auth_view") == "forgot":
        _reset_ui(admin_email, setup=False)
    else:
        _login_ui()
    return False


def change_password_widget():
    """Small expander for a logged-in admin (hidden in legacy ADMIN_PASSWORD mode)."""
    if not get_setting("ADMIN_EMAIL").strip():
        return
    with st.expander("🔑 Change password"):
        with st.form("change_pw", clear_on_submit=True):
            old = st.text_input("Current password", type="password")
            new1 = st.text_input("New password", type="password")
            new2 = st.text_input("Confirm new password", type="password")
            ok = st.form_submit_button("Change password")
        if not ok:
            return
        left = _lock_remaining()
        if left:
            st.error(f"Too many wrong attempts. Try again in {left // 60 + 1} minute(s).")
            return
        try:
            stored = _get("pw_hash", fresh=True)
        except Exception as e:
            st.error(f"Could not read the saved password: {e}")
            return
        if not (stored and verify_password(old, stored)):
            _note_failure()
            time.sleep(1)
            st.error("Current password is wrong.")
        elif password_problem(new1):
            st.error(password_problem(new1))
        elif new1 != new2:
            st.error("The two new passwords do not match.")
        else:
            try:
                token = _save_password(new1)
            except Exception as e:
                st.error(f"Could not save the password: {e}")
            else:
                st.session_state["auth_token"] = token       # this device stays logged in
                st.session_state["auth_checked"] = time.time()
                _note_success()
                st.success("✅ Password changed. All other devices were signed out.")
