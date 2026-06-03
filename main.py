#!/usr/bin/env python3
"""
Instagram Account Buyer Bot
Production-ready Telegram bot for Railway + GitHub deployment.
"""

import json
import logging
import os
import random
import re
import string
import tempfile
from datetime import datetime, timezone
import hashlib
import hmac
import base64
import struct
import time
import firebase_admin
import openpyxl
from firebase_admin import credentials, db
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Environment Variables
# ─────────────────────────────────────────────
BOT_TOKEN = os.environ["BOT_TOKEN"]
FIREBASE_CONFIG = os.environ["FIREBASE_CONFIG"]
FIREBASE_DATABASE_URL = os.environ["FIREBASE_DATABASE_URL"]

# ─────────────────────────────────────────────
# Admin IDs
# ─────────────────────────────────────────────
ADMIN_IDS = {8907284640, 8760645843}

# ─────────────────────────────────────────────
# Conversation States
# ─────────────────────────────────────────────
(
    HOME,
    TASK_MENU,
    TASK_2FA_INFO,
    TASK_2FA_STARTED,
    TASK_2FA_AWAIT_KEY,
    WITHDRAW_MENU,
    WITHDRAW_AMOUNT,
    WITHDRAW_ADDRESS,
    ADMIN_ACTS_VIEW,
) = range(9)

# ─────────────────────────────────────────────
# Firebase Initialisation
# ─────────────────────────────────────────────
def init_firebase():
    """
    Initialise Firebase using FIREBASE_CONFIG (service account JSON)
    and FIREBASE_DATABASE_URL (separate Railway env var).
    databaseURL is NOT required inside the service account JSON.
    """
    config = json.loads(FIREBASE_CONFIG)
    # Remove databaseURL if accidentally present in the service account JSON
    config.pop("databaseURL", None)
    database_url = FIREBASE_DATABASE_URL.strip()
    if not database_url:
        raise ValueError("FIREBASE_DATABASE_URL environment variable is missing or empty.")
    cred = credentials.Certificate(config)
    firebase_admin.initialize_app(cred, {"databaseURL": database_url})
    logger.info("Firebase initialised successfully.")
    
# ─────────────────────────────────────────────
# Referral Helpers
# ─────────────────────────────────────────────
def get_referral_data(user_id: int) -> dict:
    """Get referral data for a user."""
    try:
        data = db.reference(f"referrals/{user_id}").get()
        if data is None:
            data = {
                "referral_code": str(user_id),
                "referred_by": None,
                "referrals": [],  # list of user IDs who joined via this user
                "total_earned": 0.0,
                "joined_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
            }
            db.reference(f"referrals/{user_id}").set(data)
        return data
    except Exception as e:
        logger.error(f"get_referral_data({user_id}) failed: {e}")
        return {"referral_code": str(user_id), "referred_by": None, "referrals": [], "total_earned": 0.0}


def update_referral_data(user_id: int, updates: dict):
    """Update referral data for a user."""
    try:
        db.reference(f"referrals/{user_id}").update(updates)
    except Exception as e:
        logger.error(f"update_referral_data({user_id}) failed: {e}")


def add_referral(user_id: int, referrer_id: int):
    """Add a referral relationship."""
    if user_id == referrer_id:
        return False  # Can't refer yourself
    
    try:
        # Get referrer's data
        referrer_data = get_referral_data(referrer_id)
        
        # Check if user already referred by someone
        user_data = get_referral_data(user_id)
        if user_data.get("referred_by") is not None:
            return False  # Already referred by someone
        
        # Update user's referred_by
        update_referral_data(user_id, {"referred_by": referrer_id})
        
        # Add to referrer's referrals list
        referrals = referrer_data.get("referrals", [])
        if user_id not in referrals:
            referrals.append(user_id)
            update_referral_data(referrer_id, {"referrals": referrals})
        
        # Add bonus to referrer ($0.05)
        user_balance = get_user(referrer_id)
        new_balance = round(user_balance.get("balance", 0.0) + 0.05, 4)
        update_user(referrer_id, {"balance": new_balance})
        
        # Update referrer's total earned from referrals
        total_earned = referrer_data.get("total_earned", 0.0) + 0.05
        update_referral_data(referrer_id, {"total_earned": round(total_earned, 4)})
        
        logger.info(f"Referral added: {user_id} referred by {referrer_id}")
        return True
        
    except Exception as e:
        logger.error(f"add_referral({user_id}, {referrer_id}) failed: {e}")
        return False


def get_referral_stats(user_id: int) -> dict:
    """Get referral statistics for a user."""
    data = get_referral_data(user_id)
    referrals = data.get("referrals", [])
    
    # Count new referrals in last 24 hours
    now = datetime.now(timezone.utc)
    new_last_24h = 0
    
    for ref_id in referrals:
        ref_data = get_referral_data(ref_id)
        joined_at_str = ref_data.get("joined_at", "")
        if joined_at_str:
            try:
                joined_at = datetime.strptime(joined_at_str, "%Y-%m-%d %H:%M:%S UTC").replace(tzinfo=timezone.utc)
                if (now - joined_at).total_seconds() <= 24 * 3600:
                    new_last_24h += 1
            except:
                pass
    
    return {
        "total": len(referrals),
        "new_last_24h": new_last_24h,
        "total_earned": data.get("total_earned", 0.0)
    }
    
# ─────────────────────────────────────────────
# Leaderboard Helpers
# ─────────────────────────────────────────────
def get_today_stats() -> dict:
    """Get today's submissions for all users."""
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    all_subs = get_all_submissions()
    today_stats = {}
    
    for uid, subs in all_subs.items():
        if subs:
            today_count = 0
            for sub_id, sub_data in subs.items():
                sub_date = sub_data.get('datetime', '')
                if sub_date.startswith(today):
                    today_count += 1
            if today_count > 0:
                today_stats[uid] = today_count
    return today_stats


def get_leaderboard_data() -> dict:
    """Get leaderboard data from Firebase."""
    try:
        data = db.reference("leaderboard/data").get()
        return data or {}
    except Exception as e:
        logger.error(f"get_leaderboard_data failed: {e}")
        return {}


def set_leaderboard_data(data: dict):
    """Set leaderboard data in Firebase."""
    try:
        db.reference("leaderboard/data").set(data)
    except Exception as e:
        logger.error(f"set_leaderboard_data failed: {e}")


def get_leaderboard_settings() -> dict:
    """Get leaderboard settings."""
    try:
        settings = db.reference("leaderboard/settings").get()
        if settings is None:
            settings = {
                "mode": "auto",  # "auto" or "real"
                "last_update": "",
                "current_prizes": {}
            }
            db.reference("leaderboard/settings").set(settings)
        return settings
    except Exception as e:
        logger.error(f"get_leaderboard_settings failed: {e}")
        return {"mode": "auto", "last_update": "", "current_prizes": {}}


def set_leaderboard_settings(settings: dict):
    """Set leaderboard settings."""
    try:
        db.reference("leaderboard/settings").update(settings)
    except Exception as e:
        logger.error(f"set_leaderboard_settings failed: {e}")

# ─────────────────────────────────────────────
# Leaderboard Generation
# ─────────────────────────────────────────────
def generate_real_leaderboard() -> dict:
    """Generate real-time leaderboard based on today's submissions."""
    today_stats = get_today_stats()
    
    # Sort by submissions count (descending)
    sorted_users = sorted(today_stats.items(), key=lambda x: x[1], reverse=True)[:10]
    
    # Prize distribution
    prizes = [2.0, 1.0, 0.5, 0.5, 0.5, 0.2, 0.2, 0.2, 0.2, 0.1]
    
    leaderboard = {}
    for idx, (uid, count) in enumerate(sorted_users):
        if idx < len(prizes):
            prize = prizes[idx]
            leaderboard[uid] = {
                "completed": count,
                "prize": prize,
                "rank": idx + 1
            }
    return leaderboard


def generate_auto_leaderboard(increment: int = 0) -> dict:
    """Generate random auto leaderboard — counts always descending by rank."""

    prizes = [4.0, 2.0, 1.0, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]

    # Generate 10 unique fake user IDs
    used_ids = set()
    uid_list = []
    while len(uid_list) < 10:
        uid = random.randint(100000000, 999999999)
        if uid not in used_ids:
            used_ids.add(uid)
            uid_list.append(uid)

    # Generate 10 random counts then SORT descending so rank 1 always highest
    counts = [random.randint(10, 100) for _ in range(10)]
    counts.sort(reverse=True)

    # If increment mode, add small random amount to each
    if increment:
        counts = [c + random.randint(3, 10) for c in counts]

    leaderboard = {}
    for idx, (uid, count) in enumerate(zip(uid_list, counts)):
        leaderboard[str(uid)] = {
            "completed": count,
            "prize": prizes[idx],
            "rank": idx + 1,
            "masked_id": str(uid)[:4] + "***" + str(uid)[-2:]
        }

    return leaderboard


def format_leaderboard_text(leaderboard: dict, mode: str) -> str:
    """Format leaderboard for display."""
    if not leaderboard:
        return "📊 No data available for leaderboard yet."
    
    lines = [
        "🏆 **Top 10 Users Per Day**",
        "",
        "ℹ️ Results are announced every day at 01:00 (Helsinki).",
        "Leaders receive real money to their balance!",
        "",
        "🔄 Statistics update every 10 minutes.",
        "",
        "👤 ID          | ✅ Completed | 💰 Prize",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ]
    
    for rank, (uid, data) in enumerate(sorted(leaderboard.items(), key=lambda x: x[1]['rank']), 1):
        completed = data.get('completed', 0)
        prize = data.get('prize', 0)
        
        # Mask user ID
        if len(uid) > 6:
            masked = uid[:4] + "***" + uid[-2:]
        else:
            masked = uid
        
        # Format completion text
        if mode == "auto":
            if completed >= 70:
                completion_text = f"{completed} executions"
            elif completed >= 50:
                completion_text = f"{completed}+ executions"
            else:
                completion_text = f"{completed} executions"
        else:
            completion_text = f"{completed}"
        
        lines.append(f"{rank}. {masked} — {completion_text} | ${prize}")
    
    return "\n".join(lines)
    
    

# ─────────────────────────────────────────────
# Firebase Helpers
# ─────────────────────────────────────────────
def get_user(user_id: int) -> dict:
    try:
        ref = db.reference(f"users/{user_id}")
        data = ref.get()
        if data is None:
            data = {
                "balance": 0.0,
                "approved": 0,
                "in_review": 0,
                "total_submitted": 0,
            }
            ref.set(data)
        return data
    except Exception as e:
        logger.error(f"get_user({user_id}) failed: {e}")
        return {"balance": 0.0, "approved": 0, "in_review": 0, "total_submitted": 0}


def update_user(user_id: int, updates: dict):
    try:
        db.reference(f"users/{user_id}").update(updates)
    except Exception as e:
        logger.error(f"update_user({user_id}) failed: {e}")


def add_submission(user_id: int, tg_username: str, username: str, password: str, key: str):
    try:
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        sub_ref = db.reference(f"submissions/{user_id}").push()
        sub_ref.set({
            "username": username,
            "password": password,
            "key": key,
            "tg_username": tg_username,
            "user_id": str(user_id),
            "datetime": now,
        })
        user = get_user(user_id)
        update_user(user_id, {
            "in_review": user.get("in_review", 0) + 1,
            "total_submitted": user.get("total_submitted", 0) + 1,
        })
        # Rebuild and save XLSX after every new submission
        _rebuild_xlsx(user_id)
    except Exception as e:
        logger.error(f"add_submission({user_id}) failed: {e}")


def get_submissions(user_id: int) -> list:
    try:
        data = db.reference(f"submissions/{user_id}").get()
        if not data:
            return []
        return [{"id": k, **v} for k, v in data.items()]
    except Exception as e:
        logger.error(f"get_submissions({user_id}) failed: {e}")
        return []


def remove_submissions(user_id: int):
    try:
        count = len(get_submissions(user_id))
        db.reference(f"submissions/{user_id}").delete()
        # Also remove cached XLSX from Firebase storage reference
        try:
            db.reference(f"xlsx_cache/{user_id}").delete()
        except Exception:
            pass
        user = get_user(user_id)
        new_review = max(0, user.get("in_review", 0) - count)
        update_user(user_id, {"in_review": new_review})
    except Exception as e:
        logger.error(f"remove_submissions({user_id}) failed: {e}")


def create_withdrawal(user_id: int, tg_username: str, amount: float, wallet: str) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    ref = db.reference(f"withdrawals/{user_id}").push()
    fee = 0.025
    receive = round(amount - fee, 4)
    ref.set({
        "tg_username": tg_username,
        "amount": amount,
        "fee": fee,
        "receive": receive,
        "wallet": wallet,
        "status": "pending",
        "datetime": now,
    })
    return ref.key


def get_withdrawal(user_id: int, w_id: str):
    try:
        return db.reference(f"withdrawals/{user_id}/{w_id}").get()
    except Exception as e:
        logger.error(f"get_withdrawal({user_id}, {w_id}) failed: {e}")
        return None


def update_withdrawal(user_id: int, w_id: str, updates: dict):
    try:
        db.reference(f"withdrawals/{user_id}/{w_id}").update(updates)
    except Exception as e:
        logger.error(f"update_withdrawal({user_id}, {w_id}) failed: {e}")


def get_all_users() -> dict:
    try:
        data = db.reference("users").get()
        return data or {}
    except Exception as e:
        logger.error(f"get_all_users() failed: {e}")
        return {}


def get_all_submissions() -> dict:
    try:
        data = db.reference("submissions").get()
        return data or {}
    except Exception as e:
        logger.error(f"get_all_submissions() failed: {e}")
        return {}


# ─────────────────────────────────────────────
# Task Price Settings (Firebase)
# ─────────────────────────────────────────────
def get_task_price() -> float:
    """Get current task price from Firebase."""
    try:
        price = db.reference("settings/task_price").get()
        if price is None:
            price = 0.024  # default price
            db.reference("settings/task_price").set(price)
        return float(price)
    except Exception as e:
        logger.error(f"get_task_price failed: {e}")
        return 0.024


def set_task_price(price: float):
    """Set task price in Firebase."""
    try:
        db.reference("settings/task_price").set(round(price, 4))
    except Exception as e:
        logger.error(f"set_task_price failed: {e}")
        
# ─────────────────────────────────────────────
# XLSX Helpers
# ─────────────────────────────────────────────
XLSX_DIR = tempfile.gettempdir()


def _xlsx_path(user_id: int) -> str:
    return os.path.join(XLSX_DIR, f"submissions_{user_id}.xlsx")


def _rebuild_xlsx(user_id: int):
    """Rebuild and persist XLSX for a user after every submission."""
    try:
        subs = get_submissions(user_id)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Submissions"
        ws.append(["Username", "Password", "2FA Key", "TG Username", "User ID", "DateTime"])
        for row in subs:
            ws.append([
                row.get("username", ""),
                row.get("password", ""),
                row.get("key", ""),
                row.get("tg_username", ""),
                row.get("user_id", str(user_id)),
                row.get("datetime", ""),
            ])
        path = _xlsx_path(user_id)
        wb.save(path)
        logger.info(f"XLSX rebuilt for user {user_id} → {path}")
    except Exception as e:
        logger.error(f"_rebuild_xlsx({user_id}) failed: {e}")


def build_xlsx_bytes(user_id: int) -> bytes:
    """
    Return XLSX bytes. Use cached file if available; rebuild if not.
    Cleans up temp file after reading.
    """
    path = _xlsx_path(user_id)
    if not os.path.exists(path):
        _rebuild_xlsx(user_id)
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception as e:
        logger.error(f"build_xlsx_bytes({user_id}) failed: {e}")
        return b""

# ─────────────────────────────────────────────
# Utility Functions
# ─────────────────────────────────────────────
UNCOMMON_ADJECTIVES = [
    "lunar", "crimson", "azure", "phantom", "mystic", "neon", "velvet",
    "cobalt", "sable", "ivory", "scarlet", "gilded", "obsidian", "spectral",
    "vivid", "prism", "hollow", "ancient", "frosted", "runic",
]
UNCOMMON_NOUNS = [
    "wraith", "oracle", "cipher", "nexus", "specter", "herald", "vortex",
    "relic", "sigil", "golem", "comet", "prism", "mirage", "scion",
    "phantom", "herald", "bastion", "ember", "zenith", "epoch",
]

# Fix #6: Strict hex character set — no letters outside 0-9a-f
_HEX_CHARS = "0123456789abcdef"


def generate_username() -> str:
    adj = random.choice(UNCOMMON_ADJECTIVES)
    noun = random.choice(UNCOMMON_NOUNS)
    num = random.randint(100, 9999)
    return f"{adj}_{noun}_{num}"


def generate_tx_hash() -> str:
    """Generate a realistic Ethereum-style transaction hash."""
    return "0x" + "".join(random.choices(_HEX_CHARS, k=64))


# Fix #2: Strict Base32 2FA key validation
_BASE32_RE = re.compile(r"^[A-Z2-7]{16,32}$")

def validate_2fa_key(raw: str) -> tuple:
    """
    Validate 2FA key.
    Returns: (cleaned_key_without_spaces, original_key_with_spaces, error_message)
    """
    original = raw.strip()  # ইউজারের original input (স্পেস সহ)
    cleaned = original.replace(" ", "").upper()  # স্পেস রিমুভ করে চেক করার জন্য

    if len(cleaned) not in (16, 32):
        return None, None, (
            f"❌ Invalid 2FA key length. Must be 16 or 32 characters "
            f"(you entered {len(cleaned)} characters after removing spaces). Try again:"
        )

    if not _BASE32_RE.match(cleaned):
        return None, None, (
            "❌ Invalid 2FA key format. Only letters A-Z and digits 2-7 are allowed. Try again:"
        )

    return cleaned, original, None  # cleaned (স্পেস ছাড়া), original (স্পেস সহ)


# Fix #10: BEP-20 address validation
_BEP20_RE = re.compile(r"^0x[0-9a-fA-F]{40}$")


def is_valid_bep20(address: str) -> bool:
    return bool(_BEP20_RE.match(address))

# ─────────────────────────────────────────────
# Keyboards
# ─────────────────────────────────────────────
HOME_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["💰 Balance", "📋 Tasks"],
        ["📥 Withdraw", "🫟 I'm New User"],
        ["🎁 TOP"],
        ["👥 Referrals", "💝 Support"],
    ],
    resize_keyboard=True,
)

TASK_MENU_KEYBOARD = ReplyKeyboardMarkup(
    [["Inst 2FA - 0.025$"], ["🔙 Back"]],
    resize_keyboard=True,
)

TASK_START_KEYBOARD = ReplyKeyboardMarkup(
    [["✨ Start", "Cancel ❌"]],
    resize_keyboard=True,
)

WITHDRAW_MENU_KEYBOARD = ReplyKeyboardMarkup(
    [["USDT-BEP20"], ["🔙 Back"]],
    resize_keyboard=True,
)

BACK_KEYBOARD = ReplyKeyboardMarkup(
    [["🔙 Back"]],
    resize_keyboard=True,
)

# ─────────────────────────────────────────────
# /start
# ─────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    user_id = user.id
    
    # Handle referral from start parameter
    if context.args and len(context.args) > 0:
        referrer_id_str = context.args[0]
        try:
            referrer_id = int(referrer_id_str)
            if referrer_id != user_id:  # Don't refer yourself
                # Check if user already has a referrer
                user_ref_data = get_referral_data(user_id)
                if user_ref_data.get("referred_by") is None:
                    add_referral(user_id, referrer_id)
                    await update.message.reply_text(
                        "🎉 **You've been referred!** 🎉\n\n"
                        "You received a warm welcome from an existing user!\n"
                        "Complete tasks and earn money! 💰",
                        parse_mode=None
                    )
        except ValueError:
            pass
    
    try:
        get_user(user.id)
        get_referral_data(user.id)  # Initialize referral data if not exists
    except Exception as e:
        logger.error(f"cmd_start get_user failed: {e}")
    
    # ─────────────────────────────────────────────
    # NEW: Send notification to all admins
    # ─────────────────────────────────────────────
    full_name = user.full_name or user.first_name or "User"
    username = f"@{user.username}" if user.username else "No username"
    mention = f"[{full_name}](tg://user?id={user_id})"
    
    admin_notification = (
        f"🆕 **New User Started the Bot!**\n\n"
        f"👤 **Name:** {mention}\n"
        f"🆔 **User ID:** `{user_id}`\n"
        f"📱 **Username:** {username}\n"
        f"📅 **Time:** {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}\n\n"
        f"💡 Total users so far: `{len(get_all_users())}`"
    )
    
    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(
                chat_id=admin_id,
                text=admin_notification,
                parse_mode=None
            )
        except Exception as e:
            logger.warning(f"Could not notify admin {admin_id}: {e}")
    
    await update.message.reply_text(
        f"🥰 স্বাগতম, {full_name}!\n"
        "💎 কাজ শুরু করতে নিচের অপশনগুলো ব্যবহার করুন 🔽",
        reply_markup=HOME_KEYBOARD,
    )
    context.user_data.clear()
    return HOME
    
async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Get list of all user IDs with copy functionality (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    status_msg = await update.message.reply_text("📊 Fetching user list...")

    try:
        all_users = get_all_users()
        if not all_users:
            await status_msg.edit_text("❌ No users found.")
            return

        user_ids = sorted([int(uid) for uid in all_users.keys()])
        total_users = len(user_ids)

        # Create inline keyboard with user IDs as buttons (paginated)
        from telegram import InlineKeyboardButton, InlineKeyboardMarkup

        # Show first 10 users per page
        PAGE_SIZE = 10
        total_pages = (total_users + PAGE_SIZE - 1) // PAGE_SIZE

        # Store user list in context for pagination
        context.user_data["user_list"] = user_ids
        context.user_data["total_pages"] = total_pages

        # Generate first page
        await _send_user_list_page(update, context, 1)

    except Exception as e:
        logger.error(f"cmd_list failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")


async def _send_user_list_page(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int):
    """Send a specific page of user list."""
    user_ids = context.user_data.get("user_list", [])
    total_pages = context.user_data.get("total_pages", 1)
    PAGE_SIZE = 10

    if not user_ids:
        await update.message.reply_text("❌ No user list found. Use /list again.")
        return

    start_idx = (page - 1) * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    page_users = user_ids[start_idx:end_idx]

    # Build message text
    lines = [
        f"📋 **USER LIST** (Total: {len(user_ids)} users)",
        f"📄 Page {page}/{total_pages}",
        "━━━━━━━━━━━━━━━━━━━━━━",
        ""
    ]

    for idx, uid in enumerate(page_users, start=start_idx + 1):
        lines.append(f"{idx}. `{uid}`")

    lines.append("")
    lines.append("💡 **Click any button below to copy user ID**")

    message_text = "\n".join(lines)

    # Build inline keyboard
    keyboard = []
    row = []
    for uid in page_users:
        row.append(InlineKeyboardButton(str(uid), callback_data=f"copy_uid:{uid}"))
        if len(row) == 3:  # 3 buttons per row
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)

    # Add navigation buttons
    nav_buttons = []
    if page > 1:
        nav_buttons.append(InlineKeyboardButton("◀ Previous", callback_data=f"list_page:{page - 1}"))
    if page < total_pages:
        nav_buttons.append(InlineKeyboardButton("Next ▶", callback_data=f"list_page:{page + 1}"))

    if nav_buttons:
        keyboard.append(nav_buttons)
    
    # Add close button
    keyboard.append([InlineKeyboardButton("❌ Close", callback_data="list_close")])

    reply_markup = InlineKeyboardMarkup(keyboard)

    # Send or edit message
    if update.callback_query:
        await update.callback_query.edit_message_text(
            message_text,
            parse_mode=None,
            reply_markup=reply_markup
        )
        await update.callback_query.answer()
    else:
        await update.message.reply_text(
            message_text,
            parse_mode=None,
            reply_markup=reply_markup
        )


async def callback_list_handlers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle list pagination and copy callbacks."""
    query = update.callback_query
    user_id = query.from_user.id

    # Only admins can use these
    if user_id not in ADMIN_IDS:
        await query.answer("❌ Not authorised.", show_alert=True)
        return

    data = query.data

    # Handle page navigation
    if data.startswith("list_page:"):
        page = int(data.split(":")[1])
        await _send_user_list_page(update, context, page)
        return

    # Handle close
    if data == "list_close":
        await query.edit_message_text("✅ List closed.")
        await query.answer()
        return

    # Handle copy user ID
    if data.startswith("copy_uid:"):
        uid = data.split(":")[1]
        # Telegram doesn't have native copy callback, but we can show a toast
        await query.answer(f"✅ Copied: {uid}", show_alert=False)
        # Optionally send a message with the ID in code block for easy copying
        await query.message.reply_text(
            f"📋 User ID copied to clipboard:\n`{uid}`\n\n"
            f"💡 You can now use this ID with commands like:\n"
            f"`/add 0.5 {uid}`\n"
            f"`/userinfo {uid}`\n"
            f"`/msg {uid} Hello`",
            parse_mode=None
        )
        return
    
async def handle_leaderboard(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle TOP button press."""
    settings = get_leaderboard_settings()
    mode = settings.get("mode", "auto")
    
    # Get leaderboard data
    leaderboard = get_leaderboard_data()
    
    if not leaderboard:
        await update.message.reply_text(
            "📊 Leaderboard is being generated...\n"
            "Please check back later or use /ldset to generate now."
        )
        return HOME
    
    text = format_leaderboard_text(leaderboard, mode)
    await update.message.reply_text(text, parse_mode=None)
    return HOME
    
async def cmd_ldset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Set real-time leaderboard (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return
    
    status_msg = await update.message.reply_text("📊 Generating real-time leaderboard...")
    
    try:
        # Generate real leaderboard
        leaderboard = generate_real_leaderboard()
        
        # Save to Firebase
        set_leaderboard_data(leaderboard)
        
        # Update settings
        set_leaderboard_settings({
            "mode": "real",
            "last_update": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        })
        
        # Format and show
        text = format_leaderboard_text(leaderboard, "real")
        await status_msg.edit_text(text, parse_mode=None)
        
        # Also send to admin
        await update.message.reply_text(
            f"✅ Real-time leaderboard updated!\n"
            f"📅 Last update: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
        )
        
    except Exception as e:
        logger.error(f"cmd_ldset failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")


async def cmd_ldauto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate auto leaderboard (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    status_msg = await update.message.reply_text("🎲 Generating auto leaderboard...")
    
    previous = get_leaderboard_data()
    increment = 1 if previous else 0
    
    try:
        leaderboard = generate_auto_leaderboard(increment)
        set_leaderboard_data(leaderboard)
        set_leaderboard_settings({
            "mode": "auto",
            "last_update": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        })
        text = format_leaderboard_text(leaderboard, "auto")
        await status_msg.edit_text(text, parse_mode=None)
        
        if increment:
            await update.message.reply_text(
                f"✅ Auto leaderboard updated with +5-15 executions!\n"
                f"📅 Last update: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
            )
        else:
            await update.message.reply_text(
                f"✅ Auto leaderboard generated!\n"
                f"📅 Last update: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
            )
        
        # ✅ এখানে কোনো return লাগবে না (ফাংশন auto-return করে)
        
    except Exception as e:
        logger.error(f"cmd_ldauto failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")
        # ✅ এখানেও কোনো return লাগবে না
    
async def cmd_userinfo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Get detailed information about a user (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    if not context.args:
        await update.message.reply_text(
            "📝 **Usage:** `/userinfo {userid}`\n\n"
            "📌 **Example:** `/userinfo 123456789`\n\n"
            "ℹ️ Shows all user details including balance, submissions, withdrawals.",
            parse_mode=None
        )
        return

    try:
        user_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ Invalid user ID. Must be a number.")
        return

    # Send "loading" message
    status_msg = await update.message.reply_text("🔍 Fetching user information...")

    try:
        # Get user data
        data = get_user(user_id)
        
        # Get submissions count
        submissions = get_submissions(user_id)
        submission_count = len(submissions)
        
        # Get pending withdrawals count
        pending_wd = 0
        total_withdrawn = 0.0
        try:
            withdrawals = db.reference(f"withdrawals/{user_id}").get()
            if withdrawals:
                for w_id, w_data in withdrawals.items():
                    status = w_data.get("status", "")
                    amount = w_data.get("amount", 0.0)
                    if status == "pending":
                        pending_wd += 1
                    elif status == "approved":
                        total_withdrawn += amount
        except Exception as e:
            logger.warning(f"Could not fetch withdrawals for {user_id}: {e}")
        
        # Try to get user's Telegram username (if available in submissions)
        tg_username = "Not available"
        if submissions:
            for sub in submissions:
                if sub.get("tg_username"):
                    tg_username = sub.get("tg_username")
                    break
        
        # Calculate total earned
        total_earned = (data.get('approved', 0) * get_task_price())
        
        info_text = (
            f"👤 **User Information**\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🆔 **User ID:** `{user_id}`\n"
            f"📱 **TG Username:** `{tg_username}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"💰 **Current Balance:** `${data.get('balance', 0.0):.4f}`\n"
            f"💸 **Total Withdrawn:** `${total_withdrawn:.4f}`\n"
            f"💵 **Total Earned:** `${total_earned:.4f}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"✅ **Approved Tasks:** `{data.get('approved', 0)}`\n"
            f"⏳ **In Review:** `{data.get('in_review', 0)}`\n"
            f"📊 **Total Submitted:** `{data.get('total_submitted', 0)}`\n"
            f"📝 **Pending Submissions:** `{submission_count}`\n"
            f"⏰ **Pending Withdrawals:** `{pending_wd}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"💰 **Task Price:** `${get_task_price():.4f}` per task"
        )
        
        await status_msg.edit_text(info_text, parse_mode=None)
        
    except Exception as e:
        logger.error(f"cmd_userinfo failed: {e}")
        await status_msg.edit_text(f"❌ Error fetching user info: {e}")
        
async def cmd_stp(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Set task price (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    if not context.args:
        current_price = get_task_price()
        await update.message.reply_text(
            f"💰 **Current Task Price:** `${current_price:.4f}`\n\n"
            f"📝 **Usage:** `/stp 0.030`\n"
            f"📌 **Example:** `/stp 0.025`\n\n"
            f"ℹ️ This price will be used for new task submissions.",
            parse_mode=None
        )
        return

    try:
        new_price = float(context.args[0])
        if new_price <= 0:
            await update.message.reply_text("❌ Price must be greater than 0.")
            return
        
        old_price = get_task_price()
        set_task_price(new_price)
        
        await update.message.reply_text(
            f"✅ **Task price updated!**\n\n"
            f"💰 Old price: `${old_price:.4f}`\n"
            f"💰 New price: `${new_price:.4f}`\n\n"
            f"📌 New submissions will use this price.",
            parse_mode=None
        )
        
    except ValueError:
        await update.message.reply_text(
            "❌ Invalid price. Please enter a number.\n\n"
            "📝 **Example:** `/stp 0.030`"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

# ─────────────────────────────────────────────
# HOME handlers
# ─────────────────────────────────────────────
async def handle_balance(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    try:
        data = get_user(user_id)
        balance = data.get("balance", 0.0)
        approved = data.get("approved", 0)
        in_review = data.get("in_review", 0)
    except Exception as e:
        logger.error(f"handle_balance failed: {e}")
        await update.message.reply_text("⚠️ Could not fetch balance. Try again.", reply_markup=HOME_KEYBOARD)
        return HOME

    await update.message.reply_text(
        f"Your Balance 💰\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"💵 Balance: ${balance:.4f}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"✅ Total approved: {approved}\n"
        f"⏳ In Review: {in_review}",
        reply_markup=HOME_KEYBOARD,
    )
    return HOME
    
async def handle_referrals(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle Referrals button press."""
    user_id = update.effective_user.id
    username = update.effective_user.username or str(user_id)
    
    # Get referral stats
    stats = get_referral_stats(user_id)
    
    # Generate referral link
    bot_username = context.bot.username
    referral_link = f"https://t.me/{bot_username}?start={user_id}"
    
    # Get total earned from referrals
    total_earned = stats.get("total_earned", 0.0)
    
    referral_text = (
        f"👥 **Referral Program**\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📊 **Your Stats**\n"
        f"├ Total referrals: `{stats['total']}`\n"
        f"└ New in last 24h: `{stats['new_last_24h']}`\n\n"
        f"💰 **Total earned from referrals:** `${total_earned:.4f}`\n\n"
        f"🔗 **Your referral link:**\n"
        f"`{referral_link}`\n\n"
        f"ℹ️ You receive **$0.05** of the earnings of each of your referrals\n"
        f"when they join the bot and start using it!\n\n"
        f"💡 **Tip:** Share your link with friends and earn passive income!"
    )
    
    await update.message.reply_text(referral_text, parse_mode=None)
    return HOME


async def cmd_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send a message to a user by their Telegram ID (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    if len(context.args) < 2:
        await update.message.reply_text(
            "Usage:\n/msg userid message\n\nExample:\n/msg 123456789 Hello there!"
        )
        return

    try:
        user_id = int(context.args[0])
        message = " ".join(context.args[1:])

        await context.bot.send_message(chat_id=user_id, text=message)
        await update.message.reply_text(f"✅ Message sent to user {user_id}")

    except ValueError:
        await update.message.reply_text("❌ Invalid user ID. Must be a number.")
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")


async def handle_support(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "✅ Any Problem? 24/7 Support:\n"
        "• Account related issue\n"
        "• Technical support\n"
        "• Rejected problem\n"
        "• Any query\n\n"
        "📌 Admin Contact:\n"
        "👉 @Saafi_Rhman\n"
        "👉 @its_muin",
        reply_markup=HOME_KEYBOARD,
    )
    return HOME
    
async def handle_new_user(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    
    # Explicitly add user to database
    get_user(user_id)  # This creates user if not exists
    
    # Optional: Add to a separate users list for broadcasting
    try:
        db.reference("all_users").child(str(user_id)).set({
            "user_id": user_id,
            "username": update.effective_user.username,
            "joined_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        })
    except Exception as e:
        logger.error(f"Failed to add user to all_users: {e}")
    
    await update.message.reply_text(
        "🎁 2𝙁𝘼 𝙈𝙀𝙏𝙃𝙊𝘿?!\n\n"
        "𝙇𝙞𝙣𝙠 : https://t.me/c/4297868120/51\n\n"
        "💝 𝘾𝙤𝙤𝙠𝙞𝙚 𝙈𝙚𝙩𝙝𝙤𝙙?\n\n"
        "𝙇𝙞𝙣𝙠 : https://t.me/c/4297868120/43\n\n"
        "🛑 𝘼𝙡𝙨𝙤 𝙧𝙚𝙢𝙚𝙢𝙗𝙚𝙧 𝙒𝙚 𝙙𝙤𝙣'𝙩 𝙖𝙙𝙙𝙚𝙙 𝙘𝙤𝙤𝙠𝙞𝙚 𝙩𝙖𝙨𝙠 𝙞𝙣 𝙤𝙪𝙧 𝙗𝙤𝙩...",
        reply_markup=HOME_KEYBOARD,
        disable_web_page_preview=True,
    )
    return HOME

# ─────────────────────────────────────────────
# TASKS flow
# ─────────────────────────────────────────────
async def handle_tasks(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "📋 Choose a task below:",
        reply_markup=TASK_MENU_KEYBOARD,
    )
    return TASK_MENU


async def handle_task_2fa_info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "⏳ Review time: 6-12h ⏳\n\n"
        "📋 Tasks: 📱 Create Inst (2FA)\n\n"
        "📄 Description:\n"
        "In this task, you must create a new Inst acc using only a real mobile device.\n\n"
        "❗If you use your own information, your application will be REJECTED without verification.\n\n"
        "After registration:\n"
        "👉 No need to send any info\n"
        "✅ Just Send your 2fa key 🔐.\n\n"
        "⏳ Review time: 6-12h ⏳",
        reply_markup=TASK_START_KEYBOARD,
    )
    return TASK_2FA_INFO
    
async def cmd_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Get bot statistics (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    # Send loading message
    status_msg = await update.message.reply_text("📊 Fetching statistics...")

    try:
        # Get all users
        all_users = get_all_users()
        total_users = len(all_users)
        
        # Get all submissions
        all_subs = get_all_submissions()
        total_submissions = 0
        pending_review = 0
        for uid, subs in all_subs.items():
            if subs:
                sub_count = len(subs)
                total_submissions += sub_count
                pending_review += sub_count
        
        # Get approved count from users
        total_approved = 0
        total_balance = 0.0
        total_withdrawn = 0.0
        
        for uid, udata in all_users.items():
            total_approved += udata.get('approved', 0)
            total_balance += udata.get('balance', 0.0)
        
        # Get total withdrawn from withdrawals
        try:
            all_withdrawals = db.reference("withdrawals").get()
            if all_withdrawals:
                for uid, wds in all_withdrawals.items():
                    if wds:
                        for w_id, w_data in wds.items():
                            if w_data.get('status') == 'approved':
                                total_withdrawn += w_data.get('amount', 0.0)
        except Exception as e:
            logger.warning(f"Could not fetch withdrawals for stats: {e}")
        
        # Get task price
        task_price = get_task_price()
        
        # Calculate total earned (if all pending approved)
        potential_earnings = pending_review * task_price
        
        # Calculate completion rate
        completion_rate = (total_approved / total_submissions * 100) if total_submissions > 0 else 0
        
        # Get today's stats
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        today_submissions = 0
        today_withdrawals = 0
        
        for uid, subs in all_subs.items():
            if subs:
                for sub_id, sub_data in subs.items():
                    sub_date = sub_data.get('datetime', '')
                    if sub_date.startswith(today):
                        today_submissions += 1
        
        try:
            all_withdrawals = db.reference("withdrawals").get()
            if all_withdrawals:
                for uid, wds in all_withdrawals.items():
                    if wds:
                        for w_id, w_data in wds.items():
                            wd_date = w_data.get('datetime', '')
                            if wd_date.startswith(today):
                                today_withdrawals += 1
        except:
            pass
        
        # Create statistics message
        stats_text = (
            f"📊 **BOT STATISTICS**\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👥 **USERS**\n"
            f"├ Total Users: `{total_users}`\n"
            f"└ Active Users: `{len([u for u in all_users.values() if u.get('total_submitted', 0) > 0])}`\n\n"
            
            f"📝 **SUBMISSIONS**\n"
            f"├ Total: `{total_submissions}`\n"
            f"├ Today: `{today_submissions}`\n"
            f"├ Pending Review: `{pending_review}`\n"
            f"└ Approved: `{total_approved}`\n\n"
            
            f"💰 **FINANCIALS**\n"
            f"├ Task Price: `${task_price:.4f}`\n"
            f"├ Total Balance: `${total_balance:.4f}`\n"
            f"├ Total Withdrawn: `${total_withdrawn:.4f}`\n"
            f"├ Pending Payment: `${pending_review * task_price:.4f}`\n"
            f"└ Completion Rate: `{completion_rate:.1f}%`\n\n"
            
            f"📈 **PERFORMANCE**\n"
            f"├ Avg per User: `{total_submissions/total_users if total_users > 0 else 0:.1f}` tasks\n"
            f"├ Today's WD: `{today_withdrawals}` requests\n"
            f"└ Potential Earnings: `${potential_earnings:.4f}`"
        )
        
        await status_msg.edit_text(stats_text, parse_mode=None)
        
    except Exception as e:
        logger.error(f"cmd_stats failed: {e}")
        await status_msg.edit_text(f"❌ Error fetching statistics: {e}")


async def handle_task_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.pop("task_username", None)
    await update.message.reply_text(
        "❌ Task cancelled.",
        reply_markup=HOME_KEYBOARD,
    )
    return HOME


async def handle_task_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    username = generate_username()
    context.user_data["task_username"] = username

    await update.message.reply_text(
        f"👤 Username:\n<code>{username}</code>\n\n"
        f"🔓 Password:\n<code>axiex@3</code>\n\n"
        "📱 Open account with above username and password.\n"
        "Then submit account with 2FA Key below 😄",
        parse_mode="HTML",
        reply_markup=ReplyKeyboardMarkup([["Cancel ❌"]], resize_keyboard=True),
    )
    return TASK_2FA_AWAIT_KEY


async def handle_2fa_key(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = update.message.text.strip()
    user = update.effective_user

    if raw == "Cancel ❌":
        return await handle_task_cancel(update, context)

    # Fix #2: Strict Base32 validation (স্পেস সাপোর্ট সহ)
    cleaned_key, original_key, error = validate_2fa_key(raw)
    if error:
        await update.message.reply_text(error)
        return TASK_2FA_AWAIT_KEY

    # Generate 6-digit OTP from Base32 key (স্পেস ছাড়া cleaned_key ব্যবহার করব)
    try:
        def generate_totp(secret_key: str, interval: int = 30) -> str:
            """Generate 6-digit TOTP code from Base32 secret."""
            # Decode Base32 key
            key_bytes = base64.b32decode(secret_key, casefold=True)
            
            # Get current time step
            current_time = int(time.time())
            time_step = current_time // interval
            
            # Convert time step to bytes (8 bytes, big-endian)
            time_bytes = struct.pack(">Q", time_step)
            
            # Generate HMAC-SHA1
            hmac_hash = hmac.new(key_bytes, time_bytes, hashlib.sha1).digest()
            
            # Dynamic truncation
            offset = hmac_hash[-1] & 0x0F
            code_bytes = hmac_hash[offset:offset + 4]
            otp = struct.unpack(">I", code_bytes)[0] & 0x7FFFFFFF
            
            # Get last 6 digits
            otp = otp % 1000000
            
            return f"{otp:06d}"
        
        otp_code = generate_totp(cleaned_key)  # cleaned_key ব্যবহার করছি (স্পেস ছাড়া)
        
    except Exception as e:
        logger.error(f"Failed to generate OTP from 2FA key: {e}")
        await update.message.reply_text(
            "⚠️ Invalid 2FA key format. Please enter a valid Base32 key.\n\n"
            "✅ Valid example: `JBSWY3DPEHPK3PXP` or `JBSWY 3DPEH PK3PX P`\n"
            "📝 Make sure the key is 16 or 32 characters long (spaces ignored).",
            parse_mode=None
        )
        return TASK_2FA_AWAIT_KEY

    # Store submission data in context temporarily
    # সেভ করার জন্য original_key (স্পেস সহ) ব্যবহার করব
    context.user_data["pending_2fa_key"] = original_key if original_key else cleaned_key
    context.user_data["pending_username"] = context.user_data.get("task_username", generate_username())
    
    # Create reply keyboard for confirmation
    keyboard = ReplyKeyboardMarkup(
        [
            ["✅ Account Registered"],
            ["❌ Cancel"]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    
    await update.message.reply_text(
        f"🔐 **2FA Verification Code Generated!**\n\n"
        f"📱 **Your 6-digit code:** `{otp_code}`\n\n"
        f"⚠️ **Instructions:**\n"
        f"1. Open Instagram app\n"
        f"2. Enter username: `{context.user_data['pending_username']}`\n"
        f"3. Enter password: `axiex@3`\n"
        f"4. When asked for 2FA code, enter: `{otp_code}`\n\n"
        f"✅ After successful login, click **Account Registered**\n"
        f"❌ Click **Cancel** to abort this task\n\n"
        f"📌 **Note:** This code changes every 30 seconds!",
        parse_mode=None,
        reply_markup=keyboard
    )
    return TASK_2FA_STARTED


async def handle_account_registered(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle user confirming account registration."""
    user = update.effective_user
    
    # Get pending data
    key = context.user_data.get("pending_2fa_key")
    ig_username = context.user_data.get("pending_username", generate_username())
    
    if not key:
        await update.message.reply_text(
            "⚠️ Session expired. Please start the task again.",
            reply_markup=HOME_KEYBOARD
        )
        return HOME
    
    password = "axiex@3"
    tg_username = f"@{user.username}" if user.username else str(user.id)
    
    try:
        add_submission(user.id, tg_username, ig_username, password, key)
    except Exception as e:
        logger.error(f"handle_account_registered add_submission failed: {e}")
        await update.message.reply_text(
            "⚠️ Failed to save submission. Please try again.",
            reply_markup=HOME_KEYBOARD,
        )
        return HOME
    
    # Clear pending data
    context.user_data.pop("pending_2fa_key", None)
    context.user_data.pop("pending_username", None)
    context.user_data.pop("task_username", None)
    
    await update.message.reply_text(
        "✅ **Account Successfully Registered!**\n\n"
        "📋 Your submission has been sent for review.\n"
        f"⏳ Review time: 6-12 hours\n\n"
        f"💰 Upon approval, you will receive `+${get_task_price():.4f}`\n\n"
        f"Thank you for your work! 🎉",
        parse_mode=None,
        reply_markup=HOME_KEYBOARD
    )
    return HOME


async def handle_2fa_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle user cancelling the 2FA task."""
    context.user_data.pop("pending_2fa_key", None)
    context.user_data.pop("pending_username", None)
    context.user_data.pop("task_username", None)
    
    await update.message.reply_text(
        "❌ **Task Cancelled**\n\n"
        "You can start a new task anytime from the Tasks menu.",
        reply_markup=HOME_KEYBOARD
    )
    return HOME

# ─────────────────────────────────────────────
# WITHDRAW flow
# ─────────────────────────────────────────────
async def handle_withdraw(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "📥 Choose withdrawal method:",
        reply_markup=WITHDRAW_MENU_KEYBOARD,
    )
    return WITHDRAW_MENU


async def handle_withdraw_bep20(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    try:
        data = get_user(user_id)
        balance = data.get("balance", 0.0)
    except Exception as e:
        logger.error(f"handle_withdraw_bep20 failed: {e}")
        balance = 0.0
    context.user_data["withdraw_balance"] = balance

    await update.message.reply_text(
        f"💵 Your current balance: ${balance:.4f}\n\n"
        "Enter the amount you want to withdraw:\n"
        "• Minimum: $0.20\n"
        "• Fee: $0.025",
        reply_markup=BACK_KEYBOARD,
    )
    return WITHDRAW_AMOUNT


async def handle_withdraw_amount(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()

    if text == "🔙 Back":
        await update.message.reply_text("🏠 Back to home.", reply_markup=HOME_KEYBOARD)
        return HOME

    try:
        amount = float(text)
    except ValueError:
        await update.message.reply_text("❌ Please enter a valid number.")
        return WITHDRAW_AMOUNT

    balance = context.user_data.get("withdraw_balance", 0.0)

    if amount < 0.20:
        await update.message.reply_text("❌ Minimum withdrawal is $0.20.")
        return WITHDRAW_AMOUNT

    if amount > balance:
        await update.message.reply_text(
            f"❌ Insufficient balance. Your balance: ${balance:.4f}"
        )
        return WITHDRAW_AMOUNT

    context.user_data["withdraw_amount"] = amount
    await update.message.reply_text(
        "✅ Amount accepted!\n\nNow enter your USDT (BEP-20) wallet address:",
        reply_markup=BACK_KEYBOARD,
    )
    return WITHDRAW_ADDRESS


async def handle_withdraw_address(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()

    if text == "🔙 Back":
        await update.message.reply_text("🏠 Back to home.", reply_markup=HOME_KEYBOARD)
        return HOME

    # Fix #10: Validate BEP-20 address
    if not is_valid_bep20(text):
        await update.message.reply_text(
            "❌ Invalid BEP-20 wallet address.\n"
            "Must start with `0x` and be exactly 42 characters long.\n"
            "Please enter a valid address:"
        )
        return WITHDRAW_ADDRESS

    wallet = text
    amount = context.user_data.get("withdraw_amount", 0.0)
    fee = 0.025
    receive = round(amount - fee, 4)
    user = update.effective_user
    tg_username = f"@{user.username}" if user.username else str(user.id)

    try:
        w_id = create_withdrawal(user.id, tg_username, amount, wallet)
    except Exception as e:
        logger.error(f"create_withdrawal failed: {e}")
        await update.message.reply_text(
            "⚠️ Failed to create withdrawal request. Please try again.",
            reply_markup=HOME_KEYBOARD,
        )
        return HOME

    context.user_data.clear()

    await update.message.reply_text(
        f"✅ Withdrawal request created!\n\n"
        f"💳 Method: USDT (BEP-20)\n"
        f"👛 Wallet: {wallet}\n"
        f"💵 Debit amount: ${amount:.4f}\n"
        f"📉 Fee: $0.0250\n"
        f"💰 You will receive: ${receive:.4f}",
        reply_markup=HOME_KEYBOARD,
    )

    admin_text = (
        f"🔔 New Withdrawal Request\n\n"
        f"👤 User: {tg_username}\n"
        f"🆔 Chat ID: {user.id}\n"
        f"💵 Amount: ${amount:.4f}\n"
        f"📉 Fee: $0.0250\n"
        f"💰 Receive: ${receive:.4f}\n"
        f"👛 Wallet: {wallet}\n"
        f"🔑 Request ID: {w_id}"
    )
    inline_kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("Approve 👍", callback_data=f"wd_approve:{user.id}:{w_id}"),
            InlineKeyboardButton("Cancel ❌", callback_data=f"wd_cancel:{user.id}:{w_id}"),
        ]
    ])

    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(
                chat_id=admin_id,
                text=admin_text,
                reply_markup=inline_kb,
            )
        except Exception as e:
            logger.warning(f"Could not notify admin {admin_id}: {e}")

    return HOME

# ─────────────────────────────────────────────
# Back → HOME from sub-menus
# ─────────────────────────────────────────────
async def handle_back_to_home(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("🏠 Back to home.", reply_markup=HOME_KEYBOARD)
    return HOME
    
async def cmd_cast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Broadcast a message to all users (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    if len(context.args) < 1:
        await update.message.reply_text(
            "Usage:\n/cast message\n\nExample:\n/cast Hello everyone!"
        )
        return

    message = " ".join(context.args)
    
    # Send acknowledgement to admin
    status_msg = await update.message.reply_text("📢 Broadcasting message to all users...")
    
    try:
        all_users = get_all_users()
        if not all_users:
            await status_msg.edit_text("❌ No users found.")
            return
        
        success_count = 0
        fail_count = 0
        
        for user_id_str in all_users.keys():
            try:
                user_id = int(user_id_str)
                await context.bot.send_message(chat_id=user_id, text=message)
                success_count += 1
            except Exception as e:
                fail_count += 1
                logger.warning(f"Failed to send broadcast to {user_id_str}: {e}")
        
        await status_msg.edit_text(
            f"✅ Broadcast completed!\n\n"
            f"📨 Sent: {success_count} users\n"
            f"❌ Failed: {fail_count} users"
        )
        
    except Exception as e:
        await status_msg.edit_text(f"❌ Error: {e}")
        
async def cmd_acts(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """View and manage pending submissions of a user (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if not args:
        await update.message.reply_text(
            "📝 **Usage:** `/acts {userid}`\n\n"
            "📌 **Example:** `/acts 123456789`\n\n"
            "ℹ️ Shows all pending/in-review submissions of a user.\n"
            "You can approve or cancel each submission individually.",
            parse_mode=None
        )
        return

    try:
        target_id = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ Invalid user ID. Must be a number.")
        return

    # Store target user ID in context
    context.user_data["acts_target_id"] = target_id
    
    # Get pending submissions
    submissions = get_submissions(target_id)
    
    if not submissions:
        await update.message.reply_text(
            f"📭 User `{target_id}` has no pending submissions.",
            parse_mode=None
        )
        return
    
    # Filter only pending/in-review submissions (all submissions are pending by default)
    pending_subs = [sub for sub in submissions]
    
    if not pending_subs:
        await update.message.reply_text(
            f"✅ User `{target_id}` has no pending submissions.\n"
            f"All submissions have been processed.",
            parse_mode=None
        )
        return
    
    # Store all pending submissions in context
    context.user_data["acts_pending_subs"] = pending_subs
    context.user_data["acts_current_index"] = 0
    context.user_data["acts_approved_count"] = 0
    context.user_data["acts_cancelled_count"] = 0
    
    # Send first submission
    await _send_submission_for_review(update, context, 0)
    return ADMIN_ACTS_VIEW


async def _send_submission_for_review(update: Update, context: ContextTypes.DEFAULT_TYPE, index: int):
    """Send a single submission for admin review."""
    pending_subs = context.user_data.get("acts_pending_subs", [])
    target_id = context.user_data.get("acts_target_id")
    
    if not pending_subs or index >= len(pending_subs):
        # All submissions processed
        approved_count = context.user_data.get("acts_approved_count", 0)
        cancelled_count = context.user_data.get("acts_cancelled_count", 0)
        
        summary = (
            f"✅ **Review Completed!**\n\n"
            f"👤 **User ID:** `{target_id}`\n"
            f"✅ **Approved:** `{approved_count}`\n"
            f"❌ **Cancelled/Rejected:** `{cancelled_count}`\n"
            f"📊 **Total Processed:** `{approved_count + cancelled_count}`"
        )
        
        # Clean up context data
        context.user_data.pop("acts_pending_subs", None)
        context.user_data.pop("acts_target_id", None)
        context.user_data.pop("acts_current_index", None)
        context.user_data.pop("acts_approved_count", None)
        context.user_data.pop("acts_cancelled_count", None)
        
        if update.callback_query:
            await update.callback_query.edit_message_text(summary, parse_mode=None)
            await update.callback_query.answer()
        else:
            await update.message.reply_text(summary, parse_mode=None)
        return
    
    sub = pending_subs[index]
    total = len(pending_subs)
    current = index + 1
    
    # Format submission details
    submission_text = (
        f"📋 **Submission Review**\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 **User ID:** `{target_id}`\n"
        f"📊 **Progress:** `{current}/{total}`\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"🔐 **Instagram Username:**\n`{sub.get('username', 'N/A')}`\n\n"
        f"🔑 **Password:**\n`{sub.get('password', 'N/A')}`\n\n"
        f"🔢 **2FA Key:**\n`{sub.get('key', 'N/A')}`\n\n"
        f"📱 **TG Username:** `{sub.get('tg_username', 'N/A')}`\n"
        f"🕐 **Submitted:** `{sub.get('datetime', 'N/A')}`\n"
        f"🆔 **Submission ID:** `{sub.get('id', 'N/A')}`"
    )
    
    # Build inline keyboard
    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Approve", callback_data=f"acts_approve:{target_id}:{sub.get('id')}:{index}"),
            InlineKeyboardButton("❌ Cancel/Reject", callback_data=f"acts_cancel:{target_id}:{sub.get('id')}:{index}"),
        ],
        [
            InlineKeyboardButton("🏠 Exit & Home", callback_data="acts_exit"),
        ]
    ])
    
    # Send or edit message
    if update.callback_query:
        await update.callback_query.edit_message_text(
            submission_text,
            parse_mode=None,
            reply_markup=keyboard
        )
        await update.callback_query.answer()
    else:
        await update.message.reply_text(
            submission_text,
            parse_mode=None,
            reply_markup=keyboard
        )


async def callback_acts_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle approve/cancel actions for submissions."""
    query = update.callback_query
    admin_id = query.from_user.id
    
    # Only admins can use these
    if admin_id not in ADMIN_IDS:
        await query.answer("❌ Not authorised.", show_alert=True)
        return
    
    data = query.data
    
    # Handle exit
    if data == "acts_exit":
        # Clean up context
        context.user_data.pop("acts_pending_subs", None)
        context.user_data.pop("acts_target_id", None)
        context.user_data.pop("acts_current_index", None)
        context.user_data.pop("acts_approved_count", None)
        context.user_data.pop("acts_cancelled_count", None)
        
        await query.edit_message_text("🏠 **Review session ended.**\n\nYou can start a new review with `/acts {userid}`")
        await query.answer()
        return
    
    # Handle approve
    if data.startswith("acts_approve:"):
        _, target_id_str, sub_id, index_str = data.split(":")
        target_id = int(target_id_str)
        index = int(index_str)
        
        # Get submission details before removing
        pending_subs = context.user_data.get("acts_pending_subs", [])
        current_sub = None
        for sub in pending_subs:
            if sub.get('id') == sub_id:
                current_sub = sub
                break
        
        if current_sub:
            # Remove this specific submission from Firebase
            try:
                db.reference(f"submissions/{target_id}/{sub_id}").delete()
                logger.info(f"Admin {admin_id} approved submission {sub_id} for user {target_id}")
            except Exception as e:
                logger.error(f"Failed to delete submission {sub_id}: {e}")
                await query.answer("❌ Failed to approve submission.", show_alert=True)
                return
            
            # Update user's approved count and balance
            try:
                user_data = get_user(target_id)
                task_price = get_task_price()
                
                new_approved = user_data.get("approved", 0) + 1
                new_balance = round(user_data.get("balance", 0.0) + task_price, 4)
                new_in_review = max(0, user_data.get("in_review", 0) - 1)
                
                update_user(target_id, {
                    "approved": new_approved,
                    "balance": new_balance,
                    "in_review": new_in_review
                })
                
                # Send success message to user
                try:
                    await context.bot.send_message(
                        chat_id=target_id,
                        text=(
                            f"✅ **Report Approved!**\n\n"
                            f"💰 **Amount:** `+${task_price:.4f}`\n"
                            f"📝 **Comment:** Your Instagram account submission has been verified and approved.\n\n"
                            f"💵 **Current Balance:** `${new_balance:.4f}`\n\n"
                            f"Thank you for your hard work! 🎉"
                        ),
                        parse_mode=None
                    )
                except Exception as e:
                    logger.warning(f"Could not notify user {target_id}: {e}")
                
            except Exception as e:
                logger.error(f"Failed to update user {target_id}: {e}")
            
            # Update counters
            approved_count = context.user_data.get("acts_approved_count", 0) + 1
            context.user_data["acts_approved_count"] = approved_count
        
        # Remove from pending list
        new_pending = [sub for sub in pending_subs if sub.get('id') != sub_id]
        context.user_data["acts_pending_subs"] = new_pending
        
        # Send next submission or summary
        current_index = context.user_data.get("acts_current_index", 0)
        if new_pending:
            # Send next submission (same index since we removed one)
            await _send_submission_for_review(update, context, current_index)
        else:
            # All done, send summary
            await _send_submission_for_review(update, context, len(new_pending))
        
        await query.answer("✅ Submission approved!")
        return
    
    # Handle cancel
    if data.startswith("acts_cancel:"):
        _, target_id_str, sub_id, index_str = data.split(":")
        target_id = int(target_id_str)
        index = int(index_str)
        
        # Get submission details before removing
        pending_subs = context.user_data.get("acts_pending_subs", [])
        current_sub = None
        for sub in pending_subs:
            if sub.get('id') == sub_id:
                current_sub = sub
                break
        
        if current_sub:
            # Remove this specific submission from Firebase
            try:
                db.reference(f"submissions/{target_id}/{sub_id}").delete()
                logger.info(f"Admin {admin_id} cancelled submission {sub_id} for user {target_id}")
            except Exception as e:
                logger.error(f"Failed to delete submission {sub_id}: {e}")
                await query.answer("❌ Failed to cancel submission.", show_alert=True)
                return
            
            # Update user's in_review count only (no balance change)
            try:
                user_data = get_user(target_id)
                new_in_review = max(0, user_data.get("in_review", 0) - 1)
                update_user(target_id, {"in_review": new_in_review})
                
                # Send rejection message to user
                try:
                    await context.bot.send_message(
                        chat_id=target_id,
                        text=(
                            f"❌ **Report Rejected**\n\n"
                            f"📝 **Reason:** Your submitted Instagram account did not meet our quality standards or verification requirements.\n\n"
                            f"💡 **Tips for approval:**\n"
                            f"• Use a real mobile device for registration\n"
                            f"• Ensure 2FA is properly enabled\n"
                            f"• Submit valid 2FA backup codes\n\n"
                            f"📌 Please submit a new account following the guidelines.\n\n"
                            f"Need help? Contact support: @Saafi_Rhman / @its_muin"
                        ),
                        parse_mode=None
                    )
                except Exception as e:
                    logger.warning(f"Could not notify user {target_id}: {e}")
                
            except Exception as e:
                logger.error(f"Failed to update user {target_id}: {e}")
            
            # Update counters
            cancelled_count = context.user_data.get("acts_cancelled_count", 0) + 1
            context.user_data["acts_cancelled_count"] = cancelled_count
        
        # Remove from pending list
        new_pending = [sub for sub in pending_subs if sub.get('id') != sub_id]
        context.user_data["acts_pending_subs"] = new_pending
        
        # Send next submission or summary
        current_index = context.user_data.get("acts_current_index", 0)
        if new_pending:
            # Send next submission
            await _send_submission_for_review(update, context, current_index)
        else:
            # All done, send summary
            await _send_submission_for_review(update, context, len(new_pending))
        
        await query.answer("❌ Submission cancelled/rejected!")
        return

# ─────────────────────────────────────────────
# Inline Callbacks — Withdrawal Approve / Cancel
# ─────────────────────────────────────────────
async def callback_withdrawal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    admin_id = query.from_user.id

    if admin_id not in ADMIN_IDS:
        await query.answer("❌ Not authorised.", show_alert=True)
        return

    await query.answer()
    data = query.data

    try:
        action, user_id_str, w_id = data.split(":")
        user_id = int(user_id_str)
    except Exception:
        await query.edit_message_text("❌ Invalid callback data.")
        return

    try:
        wd = get_withdrawal(user_id, w_id)
    except Exception as e:
        logger.error(f"callback_withdrawal get_withdrawal failed: {e}")
        await query.edit_message_text("❌ Could not fetch withdrawal record.")
        return

    if not wd:
        await query.edit_message_text("❌ Withdrawal record not found.")
        return

    # Fix #5: Double-approval protection using Firebase atomic transaction
    wd_ref = db.reference(f"withdrawals/{user_id}/{w_id}")

    def _atomic_status_update(current_data):
        """
        Firebase transaction function.
        Only proceeds if status is still 'pending'.
        Returns None to abort if already processed.
        """
        if current_data is None:
            return None
        if current_data.get("status") != "pending":
            return None  # Abort — already processed
        current_data["status"] = action  # temporary marker; corrected below
        return current_data

    try:
        result = wd_ref.transaction(_atomic_status_update)
    except Exception as e:
        logger.error(f"Firebase transaction failed: {e}")
        await query.edit_message_text("❌ Transaction error. Please try again.")
        return

    # If transaction returned None it was aborted (already processed)
    if result is None:
        existing = get_withdrawal(user_id, w_id)
        current_status = existing.get("status", "unknown") if existing else "unknown"
        await query.answer(
            f"⚠️ Request already processed ({current_status}).",
            show_alert=True,
        )
        await query.edit_message_text(
            query.message.text + f"\n\nℹ️ Already {current_status}."
        )
        return

    if action == "wd_cancel":
        update_withdrawal(user_id, w_id, {"status": "cancelled"})
        await query.edit_message_text(
            query.message.text + "\n\n❌ Request CANCELLED by admin."
        )
        try:
            await context.bot.send_message(
                chat_id=user_id,
                text="❌ Your withdrawal request was cancelled.",
            )
        except Exception as e:
            logger.warning(f"Could not message user {user_id}: {e}")

    elif action == "wd_approve":
        wd = get_withdrawal(user_id, w_id)  # re-fetch after transaction
        amount = wd.get("amount", 0.0)
        receive = wd.get("receive", 0.0)

        try:
            user_data = get_user(user_id)
            new_balance = max(0.0, round(user_data.get("balance", 0.0) - amount, 4))
            update_user(user_id, {"balance": new_balance})
            update_withdrawal(user_id, w_id, {"status": "approved"})
        except Exception as e:
            logger.error(f"callback_withdrawal approve update failed: {e}")

        tx_hash = generate_tx_hash()

        await query.edit_message_text(
            query.message.text + "\n\n✅ Request APPROVED by admin."
        )
        try:
            await context.bot.send_message(
                chat_id=user_id,
                text=(
                    f"✅ Withdrawal successful!\n"
                    f"Amount: ${receive:.4f}\n\n"
                    f"Transaction:\n{tx_hash}"
                ),
            )
        except Exception as e:
            logger.warning(f"Could not message user {user_id}: {e}")

# ─────────────────────────────────────────────
# Admin Commands
# ─────────────────────────────────────────────
def is_admin(update: Update) -> bool:
    return update.effective_user.id in ADMIN_IDS


# Fix #4: Safe message sender that splits long text
async def _send_long_message(bot, chat_id: int, text: str, chunk_size: int = 4000):
    """Send a message, splitting into chunks if it exceeds Telegram's limit."""
    lines = text.split("\n")
    chunk = ""
    for line in lines:
        if len(chunk) + len(line) + 1 > chunk_size:
            await bot.send_message(chat_id=chat_id, text=chunk)
            chunk = line + "\n"
        else:
            chunk += line + "\n"
    if chunk.strip():
        await bot.send_message(chat_id=chat_id, text=chunk)


async def cmd_live(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    try:
        all_users = get_all_users()
        all_subs = get_all_submissions()
    except Exception as e:
        await update.message.reply_text(f"❌ Firebase error: {e}")
        return

    if not all_users:
        await update.message.reply_text("No users found.")
        return

    lines = ["📊 Live User Stats\n━━━━━━━━━━━━━━━━━━━━━━"]
    for uid, udata in all_users.items():
        sub_count = len(all_subs.get(uid, {}))
        lines.append(
            f"🆔 {uid} | Submitted: {sub_count} | Balance: ${udata.get('balance', 0.0):.4f}"
        )

    full_text = "\n".join(lines)
    # Fix #4: Split if too long
    try:
        await _send_long_message(context.bot, update.effective_chat.id, full_text)
    except Exception as e:
        logger.error(f"cmd_live send failed: {e}")
        await update.message.reply_text("❌ Failed to send stats.")


async def cmd_rcv(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if not args:
        await update.message.reply_text("Usage: /rcv {userid}")
        return

    try:
        target_id = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ Invalid user ID.")
        return

    subs = get_submissions(target_id)
    if not subs:
        await update.message.reply_text("No submissions found for that user.")
        return

    # Fix #3 + #9: Use pre-built XLSX; clean up temp file after sending
    tmp_path = None
    try:
        xlsx_data = build_xlsx_bytes(target_id)
        if not xlsx_data:
            await update.message.reply_text("❌ Could not generate XLSX.")
            return

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(xlsx_data)
            tmp_path = tmp.name

        with open(tmp_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"submissions_{target_id}.xlsx",
                caption=f"📁 Submissions for user {target_id} ({len(subs)} records)",
            )
    except Exception as e:
        logger.error(f"cmd_rcv failed: {e}")
        await update.message.reply_text(f"❌ Error: {e}")
    finally:
        # Fix #9: Always clean up temp file
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if len(args) < 2:
        await update.message.reply_text("Usage: /add {amount} {userid}")
        return

    try:
        amount = float(args[0])
        target_id = int(args[1])
    except ValueError:
        await update.message.reply_text("❌ Invalid arguments.")
        return

    try:
        user_data = get_user(target_id)
        new_balance = round(user_data.get("balance", 0.0) + amount, 4)
        update_user(target_id, {"balance": new_balance})
        await update.message.reply_text(
            f"✅ Added ${amount:.4f} to user {target_id}.\nNew balance: ${new_balance:.4f}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")


async def cmd_rm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if len(args) < 2:
        await update.message.reply_text("Usage: /rm {amount} {userid}")
        return

    try:
        amount = float(args[0])
        target_id = int(args[1])
    except ValueError:
        await update.message.reply_text("❌ Invalid arguments.")
        return

    try:
        user_data = get_user(target_id)
        new_balance = max(0.0, round(user_data.get("balance", 0.0) - amount, 4))
        update_user(target_id, {"balance": new_balance})
        await update.message.reply_text(
            f"✅ Removed ${amount:.4f} from user {target_id}.\nNew balance: ${new_balance:.4f}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")


async def cmd_rmreview(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if not args:
        await update.message.reply_text("Usage: /rmreview {userid}")
        return

    try:
        target_id = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ Invalid user ID.")
        return

    try:
        remove_submissions(target_id)
        update_user(target_id, {"in_review": 0})
        await update.message.reply_text(
            f"✅ All submissions removed for user {target_id}."
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")


async def cmd_apr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if len(args) < 2:
        await update.message.reply_text("Usage: /apr {amount} {userid}")
        return

    try:
        amount = int(args[0])
        target_id = int(args[1])
    except ValueError:
        await update.message.reply_text("❌ Invalid arguments.")
        return

    try:
        user_data = get_user(target_id)
        new_approved = user_data.get("approved", 0) + amount
        new_review = max(0, user_data.get("in_review", 0) - amount)
        update_user(target_id, {"approved": new_approved, "in_review": new_review})
        await update.message.reply_text(
            f"✅ Added {amount} to approved count for user {target_id}.\n"
            f"Total approved: {new_approved}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

# ─────────────────────────────────────────────
# Fallback
# ─────────────────────────────────────────────
async def handle_unknown(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "❓ Unknown command. Use the buttons below.",
        reply_markup=HOME_KEYBOARD,
    )
    return HOME

# ─────────────────────────────────────────────
# Error Handler
# ─────────────────────────────────────────────
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error("Exception occurred:", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "⚠️ Something went wrong. Please try again."
            )
        except Exception:
            pass

# ─────────────────────────────────────────────
# /cmd — Admin command list
# ─────────────────────────────────────────────
async def cmd_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    text = (
        "📋 BOT ADMIN COMMANDS\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "👥 USER MANAGEMENT\n"
        "/add {amount} {userid} — User er balance barhao\n"
        "/rm {amount} {userid} — User er balance kamo\n"
        "/apr {amount} {userid} — Approved count barhao\n"
        "/rmreview {userid} — User er sob submissions delete koro\n"
        "/userinfo {userid} — User er full info dekho\n"
        "/list — Sob user er list dekho (1-click copy)\n"
        "/acts {userid} — User er pending submissions review koro (approve/cancel individually)\n\n"
        "📢 MESSAGING\n"
        "/msg {userid} {text} — Specific user ke message pathao\n"
        "/cast {text} — Sob user ke broadcast koro\n\n"
        "📊 STATS & LEADERBOARD\n"
        "/live — Sob user er live stats dekho\n"
        "/stats — Bot er overall statistics\n"
        "/ldset — Real-time leaderboard generate koro\n"
        "/ldauto — Auto (fake) leaderboard generate koro\n\n"
        "📁 SUBMISSIONS\n"
        "/rcv {userid} — User er submissions Excel file hisebe nao\n\n"
        "⚙️ SETTINGS\n"
        "/stp {price} — Task er price set koro (e.g. /stp 0.025)\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "/cmd — Ei list dekho"
    )
    await update.message.reply_text(text)


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main():
    init_firebase()

    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            HOME: [
                MessageHandler(filters.Regex("^💰 Balance$"), handle_balance),
                MessageHandler(filters.Regex("^📋 Tasks$"), handle_tasks),
                MessageHandler(filters.Regex("^🎁 TOP"), handle_leaderboard),
                MessageHandler(filters.Regex("^👥 Referrals$"), handle_referrals),
                MessageHandler(filters.Regex("^📥 Withdraw$"), handle_withdraw),
                MessageHandler(filters.Regex("^🫟 I'm New User$"), handle_new_user),
                MessageHandler(filters.Regex("^💝 Support$"), handle_support),
            ],
            TASK_MENU: [
                MessageHandler(filters.Regex(r"^Inst 2FA - 0\.025\$$"), handle_task_2fa_info),
                MessageHandler(filters.Regex("^🔙 Back$"), handle_back_to_home),
            ],
            TASK_2FA_INFO: [
                MessageHandler(filters.Regex("^Cancel ❌$"), handle_task_cancel),
                MessageHandler(filters.Regex("^✨ Start$"), handle_task_start),
                MessageHandler(filters.Regex("^🔙 Back$"), handle_back_to_home),
            ],
            TASK_2FA_AWAIT_KEY: [
                MessageHandler(filters.Regex("^Cancel ❌$"), handle_task_cancel),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_2fa_key),
            ],
            TASK_2FA_STARTED: [  # নতুন স্টেট যোগ করুন
        MessageHandler(filters.Regex("^✅ Account Registered$"), handle_account_registered),
        MessageHandler(filters.Regex("^❌ Cancel$"), handle_2fa_cancel),
        MessageHandler(filters.Regex("^Cancel ❌$"), handle_2fa_cancel),
            ],
    
            WITHDRAW_MENU: [
                MessageHandler(filters.Regex("^USDT-BEP20$"), handle_withdraw_bep20),
                MessageHandler(filters.Regex("^🔙 Back$"), handle_back_to_home),
            ],
            WITHDRAW_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_withdraw_amount),
            ],
            WITHDRAW_ADDRESS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_withdraw_address),
            ],
        },
        fallbacks=[
            CommandHandler("start", cmd_start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unknown),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(callback_withdrawal, pattern="^wd_(approve|cancel):"))
    app.add_handler(CallbackQueryHandler(callback_list_handlers, pattern="^(list_page:|copy_uid:|list_close)"))
    app.add_handler(CallbackQueryHandler(callback_acts_handler, pattern="^(acts_|acts_exit)"))
    app.add_handler(CommandHandler("live", cmd_live))
    app.add_handler(CommandHandler("rcv", cmd_rcv))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("rm", cmd_rm))
    app.add_handler(CommandHandler("msg", cmd_msg))
    app.add_handler(CommandHandler("cast", cmd_cast))
    app.add_handler(CommandHandler("userinfo", cmd_userinfo))
    app.add_handler(CommandHandler("stp", cmd_stp))
    app.add_handler(CommandHandler("ldset", cmd_ldset))
    app.add_handler(CommandHandler("ldauto", cmd_ldauto))
    app.add_handler(CommandHandler("acts", cmd_acts))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(CommandHandler("rmreview", cmd_rmreview))
    app.add_handler(CommandHandler("apr", cmd_apr))
    app.add_handler(CommandHandler("cmd", cmd_cmd))
    app.add_error_handler(error_handler)

    logger.info("Bot is starting (polling)…")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
