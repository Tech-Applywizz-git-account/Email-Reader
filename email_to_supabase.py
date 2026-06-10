import os
import re
import time
import json
import logging
import requests
import html
from dotenv import load_dotenv
from supabase import create_client, Client
import msal

# ─────────────────────────────────────────────
# Load environment variables
# ─────────────────────────────────────────────
load_dotenv()

SUPABASE_URL      = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_KEY      = os.getenv("SUPABASE_SERVICE_KEY", "").strip()
CLIENT_ID         = os.getenv("AZURE_CLIENT_ID", "").strip()
CLIENT_SECRET     = os.getenv("AZURE_CLIENT_SECRET", "").strip()
TENANT_ID         = os.getenv("AZURE_TENANT_ID", "").strip()
EMAIL_USER        = os.getenv("EMAIL_USER", "").strip()
DEFAULT_PASSWORD  = os.getenv("DEFAULT_PASSWORD", "Applywizz@2026").strip()
CHECK_INTERVAL    = int(os.getenv("CHECK_INTERVAL", "300").strip())

TOKEN_CACHE_FILE  = "token_cache.json"
SCOPES            = ["Mail.Read", "User.Read"]
GRAPH_ENDPOINT    = "https://graph.microsoft.com/v1.0"

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Supabase client
# ─────────────────────────────────────────────
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


# ─────────────────────────────────────────────
# MSAL Auth — Device Code Flow (login once)
# ─────────────────────────────────────────────

def load_token_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def save_token_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token() -> str | None:
    """
    Gets a valid access token.
    - First run: opens browser for login (Device Code Flow).
    - Subsequent runs: silently refreshes from cached token.
    """
    cache = load_token_cache()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache
    )

    # If in GitHub Actions, use raw refresh token string to completely avoid JSON bugs!
    refresh_token = os.environ.get("AZURE_REFRESH_TOKEN", "").strip()
    if refresh_token:
        result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)
        if result and "access_token" in result:
            logger.info("⚡ Authenticated instantly via Refresh Token!")
            return result["access_token"]
        else:
            logger.error(f"❌ Refresh token rejected: {result}")
            exit(1)

    # Local fallback logic (uses cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_token_cache(cache)
            logger.info("🔑 Token refreshed silently.")
            return result["access_token"]

    if os.environ.get("GITHUB_ACTIONS"):
        logger.error("❌ Invalid Token Cache in GitHub Actions! Please use the BASE64 string.")
        exit(1)

    # First time — Device Code Flow (user opens browser once)
    logger.info("🌐 First-time login required. Starting Device Code Flow...")
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        logger.error(f"❌ Failed to create device flow: {flow}")
        return None

    print("\n" + "=" * 60)
    print("ONE-TIME LOGIN REQUIRED")
    print("=" * 60)
    print(f"\n>> Go to: {flow['verification_uri']}")
    print(f"   Enter code: {flow['user_code']}")
    print(f"\n   Login with: {EMAIL_USER}")
    print("\n   Waiting for you to complete login in browser...")
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_token_cache(cache)
        logger.info("✅ Login successful! Token cached for future use.")
        return result["access_token"]
    else:
        logger.error(f"❌ Login failed: {result.get('error_description')}")
        return None


# ─────────────────────────────────────────────
# Microsoft Graph — Read Emails
# ─────────────────────────────────────────────

def get_new_member_emails(access_token: str) -> list:
    """
    Fetches unread emails with 'New Member Alert' or 'You have a new community member!' in subject.
    Returns list of message dicts.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # Filter: unread only (we'll filter the subject in Python to avoid Graph API complexity errors)
    params = {
        "$filter": "isRead eq false",
        "$select": "id,subject,body,receivedDateTime",
        "$top": "50",
        "$orderby": "receivedDateTime desc"
    }

    url = f"{GRAPH_ENDPOINT}/me/messages"
    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        all_messages = response.json().get("value", [])
        
        def is_target_subject(subj: str) -> bool:
            s = subj.lower()
            return "new member alert" in s or "you have a new community member" in s

        messages = [m for m in all_messages if is_target_subject(m.get("subject", ""))]
        logger.info(f"📨 Found {len(messages)} new member alert email(s).")
        return messages
    else:
        logger.error(f"❌ Failed to fetch emails: {response.status_code} — {response.text}")
        return []


def mark_email_as_read(access_token: str, message_id: str):
    """Marks an email as read so it's not processed again."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"{GRAPH_ENDPOINT}/me/messages/{message_id}"
    requests.patch(url, headers=headers, json={"isRead": True})
    logger.info("✉️  Email marked as read.")


# ─────────────────────────────────────────────
# Parsers
# ─────────────────────────────────────────────

def extract_member_email(text: str) -> str | None:
    """Extract email from 'Email: xxx@xxx.com' pattern in body."""
    pattern = r'Email.*?([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})'
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def extract_name_from_subject(subject: str) -> str | None:
    """Extract 'Shivanandini Thummala' from 'New Member Alert: Shivanandini Thummala just joined'"""
    pattern = r'New Member Alert[:\s]+(.+?)\s+just joined'
    match = re.search(pattern, subject, re.IGNORECASE)
    return match.group(1).strip() if match else None


def html_to_text(html_content: str) -> str:
    """Strip HTML tags and decode entities to get plain text."""
    text = re.sub(r'<[^>]+>', ' ', html_content)
    text = html.unescape(text)
    return re.sub(r'\s+', ' ', text).strip()


# ─────────────────────────────────────────────
# Supabase — Add User
# ─────────────────────────────────────────────

def user_already_exists(member_email: str) -> bool:
    result = supabase.table("profiles").select("id").eq("email", member_email).execute()
    return len(result.data) > 0


def add_user_to_supabase(member_email: str, full_name: str | None = None) -> bool:
    """
    1. Creates user in auth.users via Admin API.
    2. Upserts row in public.profiles.
    """
    if user_already_exists(member_email):
        logger.info(f"⏭️  Already exists, skipping: {member_email}")
        return False

    try:
        # Step 1: Create auth user
        logger.info(f"🔐 Creating auth user: {member_email}")
        auth_res = supabase.auth.admin.create_user({
            "email": member_email,
            "password": DEFAULT_PASSWORD,
            "email_confirm": True,
            "user_metadata": {
                "full_name": full_name or "",
                "source": "Career Partner Community"
            }
        })

        if not auth_res.user:
            logger.error(f"❌ Auth user creation failed: {member_email}")
            return False

        user_id = auth_res.user.id
        logger.info(f"✅ Auth user created — ID: {user_id}")

        # Step 2: Insert into profiles
        profile = {
            "id":        user_id,
            "email":     member_email,
            "full_name": full_name or "",
            "role":      "user",
            "status":    True,
            "domain":    "Career Partner Community",
            "country":   "United States of America"
        }

        res = supabase.table("profiles").upsert(profile).execute()

        if res.data:
            logger.info(f"🎉 Profile added: {member_email} (name: {full_name})")
            return True
        else:
            logger.error(f"❌ Profile insert failed: {res}")
            return False

    except Exception as e:
        logger.error(f"💥 Error adding {member_email}: {e}")
        return False


# ─────────────────────────────────────────────
# Main Loop
# ─────────────────────────────────────────────

def check_inbox():
    """Get token, fetch unread alerts, process each one."""
    access_token = get_access_token()
    if not access_token:
        logger.error("❌ Could not get access token. Skipping this cycle.")
        return

    messages = get_new_member_emails(access_token)

    if not messages:
        logger.info("📭 No new member alerts.")
        return

    for msg in messages:
        subject   = msg.get("subject", "")
        body_html = msg.get("body", {}).get("content", "")
        body_text = html_to_text(body_html)
        msg_id    = msg.get("id")

        logger.info(f"📧 Processing: {subject}")

        full_name    = extract_name_from_subject(subject)
        member_email = extract_member_email(body_text)

        if not full_name and member_email:
            # Fallback to the email prefix if name is not found
            full_name = member_email.split('@')[0]

        if member_email:
            logger.info(f"👤 Member: {full_name} | {member_email}")
            success = add_user_to_supabase(member_email, full_name)
            if success:
                mark_email_as_read(access_token, msg_id)
        else:
            logger.warning(f"⚠️  Could not extract email from body. Skipping.")
            logger.warning(f"BODY DUMP: {repr(body_text)}")


if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("🚀 Email → Supabase Monitor STARTED (Cron Job)")
    logger.info(f"   Account  : {EMAIL_USER}")
    logger.info("=" * 60)

    check_inbox()
    
    logger.info("✅ Finished checking inbox. Exiting.")
