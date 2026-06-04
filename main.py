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
REQUIRED_CHANNEL = os.environ.get("REQUIRED_CHANNEL", "")
CHANNEL_LINK = os.environ.get("CHANNEL_LINK", "")
FIREBASE_DATABASE_URL = os.environ["FIREBASE_DATABASE_URL"]

# ─────────────────────────────────────────────
# Admin IDs
# ─────────────────────────────────────────────
ADMIN_IDS = {8907284640, 8760645843}
# Exchange rate for BDT (1 USD = 120 BDT)
USD_TO_BDT_RATE = 120.0
BOT_ENABLED = True
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
        
        # NO immediate bonus anymore! Reward comes from 8% of each approved task
        
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
  
def get_bot_state() -> dict:
    """Get bot state from Firebase."""
    try:
        state = db.reference("bot/state").get()
        if state is None:
            state = {
                "enabled": True,
                "last_updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                "updated_by": None
            }
            db.reference("bot/state").set(state)
        return state
    except Exception as e:
        logger.error(f"get_bot_state failed: {e}")
        return {"enabled": True, "last_updated": "", "updated_by": None}


def set_bot_state(enabled: bool, admin_id: int = None):
    """Set bot state in Firebase."""
    try:
        state = {
            "enabled": enabled,
            "last_updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            "updated_by": admin_id
        }
        db.reference("bot/state").set(state)
        logger.info(f"Bot state changed to {'ON' if enabled else 'OFF'} by admin {admin_id}")
    except Exception as e:
        logger.error(f"set_bot_state failed: {e}")
    
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
                "mode": "auto",  # "auto", "real", or "off"
                "last_update": "",
                "current_prizes": {},
                "enabled": True  # NEW: leaderboard on/off
            }
            db.reference("leaderboard/settings").set(settings)
        return settings
    except Exception as e:
        logger.error(f"get_leaderboard_settings failed: {e}")
        return {"mode": "auto", "last_update": "", "current_prizes": {}, "enabled": True}

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


def generate_auto_leaderboard(increment: int = 0, previous_leaderboard: dict = None) -> dict:
    """Generate random auto leaderboard — counts always descending by rank."""
    
    # If we have previous leaderboard and increment mode, increase existing users' counts
    if increment and previous_leaderboard:
        leaderboard = {}
        uid_list = []
        counts = []
        
        # Get existing users and add random increment (3-5) to each
        for uid, data in previous_leaderboard.items():
            old_count = data.get('completed', 0)
            increment_amount = random.randint(3, 5)  # 3-5 করে বাড়বে
            new_count = old_count + increment_amount
            uid_list.append(uid)
            counts.append(new_count)
        
        # Sort by new counts (descending) and reassign ranks
        sorted_pairs = sorted(zip(counts, uid_list), reverse=True)
        counts = [c for c, _ in sorted_pairs]
        uid_list = [u for _, u in sorted_pairs]
        
        prizes = [4.0, 2.0, 1.0, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]
        
        for idx, (uid, count) in enumerate(zip(uid_list, counts)):
            if idx < len(prizes):
                leaderboard[uid] = {
                    "completed": count,
                    "prize": prizes[idx],
                    "rank": idx + 1,
                    "masked_id": str(uid)[:4] + "***" + str(uid)[-2:] if len(str(uid)) > 6 else str(uid)
                }
        
        return leaderboard
    
    # First time - generate completely new fake users
    prizes = [4.0, 2.0, 1.0, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]
    
    # Generate 10 unique fake user IDs (consistent IDs so they don't change)
    used_ids = set()
    uid_list = []
    while len(uid_list) < 10:
        uid = random.randint(100000000, 999999999)
        if uid not in used_ids:
            used_ids.add(uid)
            uid_list.append(str(uid))
    
    # Generate 10 random counts (starting from 10-100)
    counts = [random.randint(10, 100) for _ in range(10)]
    counts.sort(reverse=True)
    
    leaderboard = {}
    for idx, (uid, count) in enumerate(zip(uid_list, counts)):
        leaderboard[uid] = {
            "completed": count,
            "prize": prizes[idx],
            "rank": idx + 1,
            "masked_id": str(uid)[:4] + "***" + str(uid)[-2:] if len(str(uid)) > 6 else str(uid)
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
        "🔄 Statistics update every 1PM ,6+ GTM.",
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
    [["USDT-BEP20"], ["BKASH - BDT"], ["🔙 Back"]],
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
    
    # Check if bot is enabled
    bot_state = get_bot_state()
    if not bot_state.get("enabled", True):
        await update.message.reply_text(
            "🔴 **Bot is Currently Offline** 🔴\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "🤖 **Status:** `MAINTENANCE MODE`\n"
            f"📅 **Since:** `{bot_state.get('last_updated', 'Unknown')}`\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "⚠️ The bot is temporarily disabled by the administrator.\n\n"
            "💡 **Please try again later.**\n\n"
            "📞 For urgent inquiries, contact support:\n"
            "   👤 @Saafi_Rhman\n"
            "   👤 @its_muin",
            parse_mode=None
        )
        return HOME
    
    # Check if user has joined required channel
    joined = await check_user_joined(user_id, context)
    
    if not joined and REQUIRED_CHANNEL:
        # User hasn't joined - show join message
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📲 𝙹𝚘𝚒𝚗 𝙲𝚑𝚊𝚗𝚗𝚎𝚕", url=CHANNEL_LINK)],
            [InlineKeyboardButton("✔️ 𝘾𝙝𝙚𝙘𝙠", callback_data="check_join")]
        ])
        
        await update.message.reply_text(
            f"🔒 **Access Restricted**\n\n"
            f"👋 Welcome {user.first_name or 'User'}!\n\n"
            f"⚠️ **You must join our channel to use this bot!**\n\n"
            f"📢 Click the button below to join and then press **Check**.\n\n"
            f"✅ After joining, you'll get full access to the bot.",
            parse_mode=None,
            reply_markup=keyboard
        )
        return HOME
    
    # User already joined or no channel required - normal start
    return await _start_bot(update, context)


async def _start_bot(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Actual bot start function after channel verification."""
    user = update.effective_user
    user_id = user.id
    
    # Handle referral from start parameter
    if context.args and len(context.args) > 0:
        referrer_id_str = context.args[0]
        try:
            referrer_id = int(referrer_id_str)
            if referrer_id != user_id:
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
        get_referral_data(user.id)
    except Exception as e:
        logger.error(f"cmd_start get_user failed: {e}")
    
    # Send notification to all admins
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
        
async def cmd_botoff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Turn off the entire bot (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    reason = " ".join(args) if args else "No reason provided"
    admin_id = update.effective_user.id
    
    try:
        # Update bot state to OFF
        set_bot_state(False, admin_id)
        
        await update.message.reply_text(
            f"🔴 **BOT IS NOW OFFLINE** 🔴\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🤖 **Bot Status:** `DISABLED`\n"
            f"👤 **Disabled by:** `{admin_id}`\n"
            f"📅 **Time:** `{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}`\n"
            f"📝 **Reason:** {reason}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            f"⚠️ **Users will see:**\n"
            f"   '🔴 Bot is currently offline for maintenance'\n\n"
            f"✅ **To turn back ON, use:** `/boton`",
            parse_mode=None
        )
        
    except Exception as e:
        logger.error(f"cmd_botoff failed: {e}")
        await update.message.reply_text(f"❌ Error: {e}")
        
async def cmd_boton(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Turn on the entire bot (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    admin_id = update.effective_user.id
    
    try:
        # Update bot state to ON
        set_bot_state(True, admin_id)
        
        await update.message.reply_text(
            f"🟢 **BOT IS NOW ONLINE** 🟢\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🤖 **Bot Status:** `ENABLED`\n"
            f"👤 **Enabled by:** `{admin_id}`\n"
            f"📅 **Time:** `{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            f"✅ **Bot is now fully operational!**\n"
            f"📊 Users can now use all commands.",
            parse_mode=None
        )
        
    except Exception as e:
        logger.error(f"cmd_boton failed: {e}")
        await update.message.reply_text(f"❌ Error: {e}")
        
async def handle_task_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.pop("task_username", None)
    await update.message.reply_text(
        "❌ Task cancelled.",
        reply_markup=HOME_KEYBOARD,
    )
    return HOME
        
async def handle_withdraw_bkash(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle BKASH withdrawal method."""
    user_id = update.effective_user.id
    try:
        data = get_user(user_id)
        balance = data.get("balance", 0.0)
    except Exception as e:
        logger.error(f"handle_withdraw_bkash failed: {e}")
        balance = 0.0
    
    # Calculate BDT amount (silent - not shown to user in this message)
    bdt_balance = balance * USD_TO_BDT_RATE
    
    context.user_data["withdraw_balance"] = balance
    context.user_data["withdraw_method"] = "bkash"
    
    await update.message.reply_text(
        f"💵 **Your Balance:** `${balance:.4f}` USD\n"
        f"💰 **≈ {bdt_balance:.2f} BDT**\n\n"
        f"📱 **BKASH Withdrawal**\n"
        f"• Minimum withdrawal: $0.20 USD (≈ {0.20 * USD_TO_BDT_RATE:.2f} BDT)\n"
        f"• Fee: $0.025 USD (≈ {0.025 * USD_TO_BDT_RATE:.2f} BDT)\n\n"
        f"📝 Enter the amount in **USD** you want to withdraw:\n"
        f"Example: `0.50`\n\n"
        f"💡 You will receive BDT equivalent at rate 1 USD = {USD_TO_BDT_RATE} BDT",
        parse_mode=None,
        reply_markup=BACK_KEYBOARD,
    )
    return WITHDRAW_AMOUNT
    
async def handle_leaderboard(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle TOP button press."""
    settings = get_leaderboard_settings()
    
    # Check if leaderboard is enabled
    if not settings.get("enabled", True):
        await update.message.reply_text(
            "🔴 **Leaderboard is currently OFF**\n\n"
            "📊 Leaderboard has been disabled by admin.\n\n"
            "💡 Please check back later."
        )
        return HOME
    
    mode = settings.get("mode", "auto")
    
    # Get leaderboard data
    leaderboard = get_leaderboard_data()
    
    if not leaderboard:
        await update.message.reply_text(
            "📊 Leaderboard is being generated...\n"
            "Please check back later or contact admin."
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
    increment = 1 if previous and previous.get("mode") == "auto" else 0
    
    try:
        # Check if previous leaderboard is auto mode
        settings = get_leaderboard_settings()
        is_auto_mode = settings.get("mode") == "auto"
        
        if is_auto_mode and previous:
            # Increment existing fake users' counts by 3-5
            leaderboard = generate_auto_leaderboard(increment=1, previous_leaderboard=previous)
            await update.message.reply_text(
                f"📈 **Leaderboard Updated!**\n\n"
                f"✅ Each user's completed tasks increased by 3-5!\n"
                f"📅 Last update: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
            )
        else:
            # First time - generate new fake users
            leaderboard = generate_auto_leaderboard(increment=0, previous_leaderboard=None)
            await update.message.reply_text(
                f"🎲 **Auto Leaderboard Generated!**\n\n"
                f"✅ 10 fake users added to leaderboard.\n"
                f"📅 Last update: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
            )
        
        # Save to Firebase
        set_leaderboard_data(leaderboard)
        set_leaderboard_settings({
            "mode": "auto",
            "last_update": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            "enabled": True  # Auto-enable when generating
        })
        
        text = format_leaderboard_text(leaderboard, "auto")
        await status_msg.edit_text(text, parse_mode=None)
        
    except Exception as e:
        logger.error(f"cmd_ldauto failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")
    
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
        
async def handle_account_registered(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle user confirming account registration."""
    user = update.effective_user
    
    key = context.user_data.get("pending_2fa_key")
    ig_username = context.user_data.get("pending_username", generate_username())
    
    if not key:
        await update.message.reply_text(
            "⚠️ Session expired. Please start the task again.",
            reply_markup=HOME_KEYBOARD
        )
        return HOME
    
    password = "axiex@5"
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
    context.user_data.pop("pending_otp_code", None)
    
    await update.message.reply_text(
        f"✅ **Account Successfully Registered!**\n\n"
        f"📋 Your submission has been sent for review.\n"
        f"⏳ Review time: 6-12 hours\n\n"
        f"💰 Upon approval, you will receive `+${get_task_price():.4f}`\n\n"
        f"Thank you! 🎉",
        parse_mode=None,
        reply_markup=HOME_KEYBOARD
    )
    return HOME
    
async def check_bot_enabled(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Check if bot is enabled, send message if disabled."""
    bot_state = get_bot_state()
    if not bot_state.get("enabled", True):
        # Only respond to admins
        if update.effective_user and update.effective_user.id in ADMIN_IDS:
            return True
        
        if update.message:
            await update.message.reply_text(
                "🔴 **Bot is Currently Offline** 🔴\n\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                "🤖 **Status:** `MAINTENANCE MODE`\n"
                f"📅 **Since:** `{bot_state.get('last_updated', 'Unknown')}`\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
                "⚠️ The bot is temporarily disabled by the administrator.\n\n"
                "💡 **Please try again later.**\n\n"
                "📞 For urgent inquiries, contact support:\n"
                "   👤 @Saafi_Rhman\n"
                "   👤 @its_muin",
                parse_mode=None
            )
        return False
    return True
    
async def callback_check_join(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle check join button callback."""
    query = update.callback_query
    user_id = query.from_user.id
    
    await query.answer()
    
    # Check if user has joined
    joined = await check_user_joined(user_id, context)
    
    if joined:
        # User joined - start the bot
        await query.edit_message_text(
            "✅ **Verification Successful!**\n\n"
            "Thank you for joining our channel!\n"
            "Now you can use the bot. 🎉"
        )
        # Send start message directly (since update.message is None in callback)
        user = query.from_user
        user_id = user.id
        full_name = user.full_name or user.first_name or "User"
        username = f"@{user.username}" if user.username else "No username"
        mention = f"[{full_name}](tg://user?id={user_id})"

        try:
            get_user(user_id)
            get_referral_data(user_id)
        except Exception as e:
            logger.error(f"callback_check_join get_user failed: {e}")

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

        await context.bot.send_message(
            chat_id=user_id,
            text=f"🥰 স্বাগতম, {full_name}!\n"
                 "💎 কাজ শুরু করতে নিচের অপশনগুলো ব্যবহার করুন 🔽",
            reply_markup=HOME_KEYBOARD,
        )
        context.user_data.clear()
    else:
        # User hasn't joined yet
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📲 𝙹𝚘𝚒𝚗 𝙲𝚑𝚊𝚗𝚗𝚎𝚕", url=CHANNEL_LINK)],
            [InlineKeyboardButton("✔️ 𝘾𝙝𝙚𝙘𝙠 𝘼𝙜𝙖𝙞𝙣", callback_data="check_join")]
        ])
        try:
            await query.edit_message_text(
                "❌ **Not Joined Yet!**\n\n"
                "You haven't joined our channel.\n\n"
                "📢 Please click the button below to join, then press **Check Again**.\n\n"
                "⚠️ Without joining, you cannot use this bot.",
                reply_markup=keyboard
            )
        except Exception as e:
            if "Message is not modified" not in str(e):
                raise
    
async def handle_2fa_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle user cancelling the 2FA task."""
    context.user_data.pop("pending_2fa_key", None)
    context.user_data.pop("pending_username", None)
    context.user_data.pop("task_username", None)
    context.user_data.pop("pending_otp_code", None)
    
    await update.message.reply_text(
        "❌ **Task Cancelled**\n\nYou can start a new task anytime.",
        reply_markup=HOME_KEYBOARD
    )
    return HOME
    
async def check_user_joined(user_id: int, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Check if user has joined the required channel."""
    if not REQUIRED_CHANNEL:
        return True  # No channel configured
    
    try:
        chat_member = await context.bot.get_chat_member(chat_id=REQUIRED_CHANNEL, user_id=user_id)
        return chat_member.status in ["member", "administrator", "creator"]
    except Exception as e:
        logger.error(f"Check join failed for {user_id}: {e}")
        return False
        
async def cmd_ldoff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Turn off leaderboard (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    try:
        # Update settings to disable leaderboard
        set_leaderboard_settings({
            "enabled": False,
            "last_update": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        })
        
        await update.message.reply_text(
            "🔴 **Leaderboard is now OFF!**\n\n"
            "📊 Users will see: 'Leaderboard is currently disabled'\n\n"
            "💡 To turn it back ON, use: `/ldauto` or `/ldset`\n"
            "   (These commands will automatically enable it)"
        )
        
    except Exception as e:
        logger.error(f"cmd_ldoff failed: {e}")
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
    
    # Get total earned from referrals (8% of referred users' approved tasks)
    total_earned = stats.get("total_earned", 0.0)
    
    # Get current task price
    task_price = get_task_price()
    referral_percentage = 8  # 8%
    
    referral_text = (
        f"👥 **Referral Program**\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📊 **Your Stats**\n"
        f"├ Total referrals: `{stats['total']}`\n"
        f"└ New in last 24h: `{stats['new_last_24h']}`\n\n"
        f"💰 **Total earned from referrals:** `${total_earned:.6f}`\n\n"
        f"🎁 **How it works:**\n"
        f"• When your referral completes a task,\n"
        f"• You get **{referral_percentage}%** of their reward!\n"
        f"• Current task reward: `${task_price:.4f}`\n"
        f"• You get: `${task_price * 0.08:.6f}` per approved task\n\n"
        f"🔗 **Your referral link:**\n"
        f"`{referral_link}`\n\n"
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
    
def approve_submission(user_id: int, submission_id: str, task_price: float):
    """Approve a submission and give reward to user and referrer."""
    try:
        # Get submission details
        sub_ref = db.reference(f"submissions/{user_id}/{submission_id}")
        submission = sub_ref.get()
        
        if not submission:
            return False
        
        # Update user's approved count and balance
        user_data = get_user(user_id)
        new_approved = user_data.get("approved", 0) + 1
        new_balance = round(user_data.get("balance", 0.0) + task_price, 4)
        new_in_review = max(0, user_data.get("in_review", 0) - 1)
        
        update_user(user_id, {
            "approved": new_approved,
            "balance": new_balance,
            "in_review": new_in_review
        })
        
        # ─────────────────────────────────────────────
        # NEW: Give 8% referral reward to referrer
        # ─────────────────────────────────────────────
        referral_data = get_referral_data(user_id)
        referrer_id = referral_data.get("referred_by")
        
        if referrer_id:
            referral_percentage = 0.08  # 8%
            referral_reward = round(task_price * referral_percentage, 6)
            
            if referral_reward > 0:
                # Update referrer's balance
                referrer_data = get_user(referrer_id)
                new_referrer_balance = round(referrer_data.get("balance", 0.0) + referral_reward, 6)
                update_user(referrer_id, {"balance": new_referrer_balance})
                
                # Update referrer's total earned from referrals
                referrer_ref_data = get_referral_data(referrer_id)
                new_total_earned = round(referrer_ref_data.get("total_earned", 0.0) + referral_reward, 6)
                update_referral_data(referrer_id, {"total_earned": new_total_earned})
                
                logger.info(f"Referral reward {referral_reward} given to {referrer_id} for user {user_id}'s approval")
        
        # Delete the submission after approval
        sub_ref.delete()
        
        return True
        
    except Exception as e:
        logger.error(f"approve_submission failed: {e}")
        return False

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
        # Get bot state
        bot_state = get_bot_state()
        bot_status = "🟢 ONLINE" if bot_state.get("enabled", True) else "🔴 OFFLINE"
        
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
        
        # Create statistics message with bot status
        stats_text = (
            f"📊 **BOT STATISTICS**\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🤖 **Bot Status:** {bot_status}\n"
            f"📅 **Last State Change:** `{bot_state.get('last_updated', 'Never')}`\n"
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


async def handle_task_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    username = generate_username()
    context.user_data["task_username"] = username

    await update.message.reply_text(
        f"👤 Username:\n<code>{username}</code>\n\n"
        f"🔓 Password:\n<code>axiex@5</code>\n\n"
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
        
        otp_code = generate_totp(cleaned_key)
        
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
    context.user_data["pending_2fa_key"] = original_key if original_key else cleaned_key
    context.user_data["pending_username"] = context.user_data.get("task_username", generate_username())
    context.user_data["pending_otp_code"] = otp_code
    
    # Create INLINE keyboard for copy (1-click copy)
    inline_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Copy 6-Digit Code", callback_data=f"copy_otp:{otp_code}")]
    ])
    
    # Create REPLY keyboard for confirm/cancel
    reply_keyboard = ReplyKeyboardMarkup(
        [
            ["✅ Account Registered"],
            ["❌ Cancel Task"]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    
    await update.message.reply_text(
        f"🔐 **Your 6-Digit Code:**\n"
        f"<code>{otp_code}</code>\n\n"
        f"💡 **Click the 'Copy Code' button below to copy!**\n\n"
        f"📱 **Instructions:**\n"
        f"1. Open Instagram app\n"
        f"2. Enter username: <code>{context.user_data['pending_username']}</code>\n"
        f"3. Enter password: <code>axiex@5</code>\n"
        f"4. Enter this code\n\n"
        f"✅ After login, click 'Account Registered'\n"
        f"❌ Click 'Cancel Task' to abort",
        parse_mode="HTML",
        reply_markup=reply_keyboard
    )
    
    # Send inline keyboard separately (for copy button)
    await update.message.reply_text(
        "👇 **Click below to copy the code:**",
        reply_markup=inline_keyboard
    )
    
    return TASK_2FA_STARTED

# ─────────────────────────────────────────────
# WITHDRAW flow
# ─────────────────────────────────────────────
async def handle_withdraw(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "📥 **Choose withdrawal method:**\n\n"
        "💰 **USDT-BEP20** - Crypto withdrawal\n"
        "📱 **BKASH - BDT** - Direct mobile banking\n\n"
        "💡 Minimum withdrawal: $0.20 USD",
        parse_mode=None,
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
        await update.message.reply_text("❌ Minimum withdrawal is $0.20 USD.")
        return WITHDRAW_AMOUNT

    if amount > balance:
        await update.message.reply_text(
            f"❌ Insufficient balance. Your balance: ${balance:.4f} USD"
        )
        return WITHDRAW_AMOUNT

    context.user_data["withdraw_amount"] = amount
    
    # Check which method was selected
    method = context.user_data.get("withdraw_method", "usdt")
    
    if method == "bkash":
        await update.message.reply_text(
            "✅ Amount accepted!\n\n"
            "📱 Now enter your **BKASH account number** (must be a valid BKASH number):\n\n"
            "📌 Example: `01XXXXXXXXX`\n\n"
            "⚠️ Make sure the number is correct to receive payment.",
            reply_markup=BACK_KEYBOARD,
        )
        return WITHDRAW_ADDRESS
    else:
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

    method = context.user_data.get("withdraw_method", "usdt")
    amount = context.user_data.get("withdraw_amount", 0.0)
    fee = 0.025
    receive_usd = round(amount - fee, 4)
    
    # Calculate receive amount in BDT if BKASH method
    if method == "bkash":
        # Validate BKASH number (Bangladeshi mobile number)
        bkash_pattern = re.compile(r"^01[3-9]\d{8}$")
        if not bkash_pattern.match(text):
            await update.message.reply_text(
                "❌ Invalid BKASH number.\n"
                "Must be a valid Bangladeshi mobile number.\n"
                "Example: `017XXXXXXXX` or `019XXXXXXXX`\n\n"
                "Please enter a valid BKASH number:"
            )
            return WITHDRAW_ADDRESS
        wallet = text
        receive_bdt = round(receive_usd * USD_TO_BDT_RATE, 2)
        receive_display = f"{receive_bdt:.2f} BDT (${receive_usd:.4f} USD)"
    else:
        # USDT method - validate BEP-20 address
        if not is_valid_bep20(text):
            await update.message.reply_text(
                "❌ Invalid BEP-20 wallet address.\n"
                "Must start with `0x` and be exactly 42 characters long.\n"
                "Please enter a valid address:"
            )
            return WITHDRAW_ADDRESS
        wallet = text
        receive_display = f"${receive_usd:.4f}"

    user = update.effective_user
    tg_username = f"@{user.username}" if user.username else str(user.id)

    try:
        # Store method in withdrawal record
        w_id = create_withdrawal(user.id, tg_username, amount, wallet)
        # Update withdrawal record with method
        db.reference(f"withdrawals/{user.id}/{w_id}").update({"method": method})
    except Exception as e:
        logger.error(f"create_withdrawal failed: {e}")
        await update.message.reply_text(
            "⚠️ Failed to create withdrawal request. Please try again.",
            reply_markup=HOME_KEYBOARD,
        )
        return HOME

    context.user_data.clear()

    # Send confirmation to user
    if method == "bkash":
        await update.message.reply_text(
            f"✅ **Withdrawal request created!**\n\n"
            f"💳 **Method:** BKASH - BDT\n"
            f"📱 **BKASH Number:** `{wallet}`\n"
            f"💵 **Debit amount:** `${amount:.4f}` USD\n"
            f"📉 **Fee:** $0.0250 USD\n"
            f"💰 **You will receive:** {receive_display}\n\n"
            f"⏳ **Processing time:** 24-48 hours\n"
            f"📞 Make sure your BKASH number is active!",
            parse_mode=None,
            reply_markup=HOME_KEYBOARD,
        )
    else:
        await update.message.reply_text(
            f"✅ **Withdrawal request created!**\n\n"
            f"💳 **Method:** USDT (BEP-20)\n"
            f"👛 **Wallet:** `{wallet}`\n"
            f"💵 **Debit amount:** `${amount:.4f}`\n"
            f"📉 **Fee:** $0.0250\n"
            f"💰 **You will receive:** {receive_display}",
            parse_mode=None,
            reply_markup=HOME_KEYBOARD,
        )

    # Admin notification
    if method == "bkash":
        admin_text = (
            f"🔔 **New Withdrawal Request**\n\n"
            f"👤 User: {tg_username}\n"
            f"🆔 Chat ID: {user.id}\n"
            f"💳 Method: BKASH - BDT\n"
            f"📱 BKASH: `{wallet}`\n"
            f"💵 Amount: ${amount:.4f} USD\n"
            f"📉 Fee: $0.0250 USD\n"
            f"💰 Receive: {receive_display}\n"
            f"🔑 Request ID: {w_id}"
        )
    else:
        admin_text = (
            f"🔔 **New Withdrawal Request**\n\n"
            f"👤 User: {tg_username}\n"
            f"🆔 Chat ID: {user.id}\n"
            f"💳 Method: USDT (BEP-20)\n"
            f"💵 Amount: ${amount:.4f}\n"
            f"📉 Fee: $0.0250\n"
            f"💰 Receive: ${receive_usd:.4f}\n"
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
                parse_mode=None,
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
    """Send submission with COPY buttons for username, password, 2FA key."""
    pending_subs = context.user_data.get("acts_pending_subs", [])
    target_id = context.user_data.get("acts_target_id")
    
    if not pending_subs or index >= len(pending_subs):
        approved_count = context.user_data.get("acts_approved_count", 0)
        cancelled_count = context.user_data.get("acts_cancelled_count", 0)
        
        summary = f"✅ **Review Completed!**\n\n👤 **User ID:** `{target_id}`\n✅ **Approved:** `{approved_count}`\n❌ **Cancelled:** `{cancelled_count}`"
        
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
    
    ig_username = sub.get('username', 'N/A')
    password = sub.get('password', 'N/A')
    twofa_key = sub.get('key', 'N/A')
    
    submission_text = (
        f"📋 **Submission Review**\n━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 **User ID:** `{target_id}`\n📊 **Progress:** `{current}/{total}`\n━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"🔐 **Instagram Username:**\n<code>{ig_username}</code>\n\n"
        f"🔑 **Password:**\n<code>{password}</code>\n\n"
        f"🔢 **2FA Key:**\n<code>{twofa_key}</code>\n\n"
        f"📱 **TG Username:** `{sub.get('tg_username', 'N/A')}`\n"
        f"🕐 **Submitted:** `{sub.get('datetime', 'N/A')}`"
    )
    
    keyboard = [
        [InlineKeyboardButton("📋 Copy Username", callback_data=f"acts_copy_username:{target_id}:{sub.get('id')}:{index}")],
        [InlineKeyboardButton("🔑 Copy Password", callback_data=f"acts_copy_password:{target_id}:{sub.get('id')}:{index}")],
        [InlineKeyboardButton("🔢 Copy 2FA Key", callback_data=f"acts_copy_2fa:{target_id}:{sub.get('id')}:{index}")],
        [InlineKeyboardButton("✅ Approve", callback_data=f"acts_approve:{target_id}:{sub.get('id')}:{index}"),
         InlineKeyboardButton("❌ Cancel", callback_data=f"acts_cancel:{target_id}:{sub.get('id')}:{index}")],
        [InlineKeyboardButton("🏠 Exit", callback_data="acts_exit")],
    ]
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    if update.callback_query:
        await update.callback_query.edit_message_text(submission_text, parse_mode="HTML", reply_markup=reply_markup)
        await update.callback_query.answer()
    else:
        await update.message.reply_text(submission_text, parse_mode="HTML", reply_markup=reply_markup)
        
async def callback_acts_copy_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle copy for username/password/2FA key."""
    query = update.callback_query
    admin_id = query.from_user.id
    
    if admin_id not in ADMIN_IDS:
        await query.answer("❌ Not authorised.", show_alert=True)
        return
    
    data = query.data
    pending_subs = context.user_data.get("acts_pending_subs", [])
    
    if data.startswith("acts_copy_username:"):
        _, _, sub_id, _ = data.split(":")
        for sub in pending_subs:
            if sub.get('id') == sub_id:
                await query.answer("✅ Copied!", show_alert=False)
                await query.message.reply_text(f"📋 **Username:**\n<code>{sub.get('username', 'N/A')}</code>\n\n💡 Press and hold to copy!", parse_mode="HTML")
                return
    
    elif data.startswith("acts_copy_password:"):
        _, _, sub_id, _ = data.split(":")
        for sub in pending_subs:
            if sub.get('id') == sub_id:
                await query.answer("✅ Copied!", show_alert=False)
                await query.message.reply_text(f"📋 **Password:**\n<code>{sub.get('password', 'N/A')}</code>\n\n💡 Press and hold to copy!", parse_mode="HTML")
                return
    
    elif data.startswith("acts_copy_2fa:"):
        _, _, sub_id, _ = data.split(":")
        for sub in pending_subs:
            if sub.get('id') == sub_id:
                await query.answer("✅ Copied!", show_alert=False)
                await query.message.reply_text(f"📋 **2FA Key:**\n<code>{sub.get('key', 'N/A')}</code>\n\n💡 Press and hold to copy!", parse_mode="HTML")
                return
                
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
            task_price = get_task_price()
            
            # Approve submission and give reward (including referral 8%)
            success = approve_submission(target_id, sub_id, task_price)
            
            if not success:
                await query.answer("❌ Failed to approve submission.", show_alert=True)
                return
            
            logger.info(f"Admin {admin_id} approved submission {sub_id} for user {target_id}")
            
            # Send success message to user
            try:
                user_data = get_user(target_id)
                new_balance = user_data.get("balance", 0.0)
                
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
            
            # Update counters
            approved_count = context.user_data.get("acts_approved_count", 0) + 1
            context.user_data["acts_approved_count"] = approved_count
        
        # Remove from pending list
        new_pending = [sub for sub in pending_subs if sub.get('id') != sub_id]
        context.user_data["acts_pending_subs"] = new_pending
        
        # Send next submission or summary
        current_index = context.user_data.get("acts_current_index", 0)
        if new_pending:
            await _send_submission_for_review(update, context, current_index)
        else:
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
            await _send_submission_for_review(update, context, current_index)
        else:
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
    """Get list of users who have submitted at least 1 account (admin only)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    status_msg = await update.message.reply_text("📊 Fetching active users...")

    try:
        all_users = get_all_users()
        all_subs = get_all_submissions()
        
        if not all_users:
            await status_msg.edit_text("❌ No users found.")
            return
        
        # Create list of users who have at least 1 submission
        active_users = []
        for uid, udata in all_users.items():
            sub_count = len(all_subs.get(uid, {}))
            if sub_count > 0:  # শুধু যারা সাবমিট করেছে
                active_users.append({
                    "user_id": int(uid),
                    "submissions": sub_count,
                    "balance": udata.get('balance', 0.0)
                })
        
        # Sort by submissions count (highest first)
        active_users.sort(key=lambda x: x['submissions'], reverse=True)
        
        if not active_users:
            await status_msg.edit_text("📭 No active users found. (Users who submitted at least 1 account)")
            return
        
        total_users = len(active_users)
        
        # Pagination setup
        PAGE_SIZE = 10
        total_pages = (total_users + PAGE_SIZE - 1) // PAGE_SIZE
        
        # Store in context for pagination
        context.user_data["live_users"] = active_users
        context.user_data["live_total_pages"] = total_pages
        
        # Send first page
        await _send_live_page(update, context, 1)
        
    except Exception as e:
        logger.error(f"cmd_live failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")


async def _send_live_page(update: Update, context: ContextTypes.DEFAULT_TYPE, page: int):
    """Send a specific page of live users list with 1-click copy."""
    active_users = context.user_data.get("live_users", [])
    total_pages = context.user_data.get("live_total_pages", 1)
    PAGE_SIZE = 10
    
    if not active_users:
        await update.message.reply_text("❌ No active users found.")
        return
    
    start_idx = (page - 1) * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    page_users = active_users[start_idx:end_idx]
    
    lines = [
        f"📊 **ACTIVE USERS**",
        f"📄 Page {page}/{total_pages}",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        ""
    ]
    
    for idx, user in enumerate(page_users, start=start_idx + 1):
        lines.append(
            f"`{idx}.` 🆔 **User ID:** `{user['user_id']}`\n"
            f"   📝 **Submissions:** `{user['submissions']}`\n"
            f"   💰 **Balance:** `${user['balance']:.4f}`\n"
        )
    
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("💡 **Click a button below to copy User ID**")
    
    message_text = "\n".join(lines)
    
    # Each user gets their own COPY button
    keyboard = []
    for user in page_users:
        keyboard.append([
            InlineKeyboardButton(
                f"📋 Copy ID: {user['user_id']}", 
                callback_data=f"live_copy_uid:{user['user_id']}"
            )
        ])
    
    # Navigation buttons
    nav_buttons = []
    if page > 1:
        nav_buttons.append(InlineKeyboardButton("◀ Previous", callback_data=f"live_page:{page - 1}"))
    if page < total_pages:
        nav_buttons.append(InlineKeyboardButton("Next ▶", callback_data=f"live_page:{page + 1}"))
    
    if nav_buttons:
        keyboard.append(nav_buttons)
    keyboard.append([InlineKeyboardButton("❌ Close", callback_data="live_close")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    if update.callback_query:
        await update.callback_query.edit_message_text(message_text, parse_mode=None, reply_markup=reply_markup)
        await update.callback_query.answer()
    else:
        await update.message.reply_text(message_text, parse_mode=None, reply_markup=reply_markup)


async def callback_live_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle live list - 1-click copy."""
    query = update.callback_query
    admin_id = query.from_user.id
    
    if admin_id not in ADMIN_IDS:
        await query.answer("❌ Not authorised.", show_alert=True)
        return
    
    data = query.data
    
    if data.startswith("live_page:"):
        page = int(data.split(":")[1])
        await _send_live_page(update, context, page)
        return
    
    if data == "live_close":
        await query.edit_message_text("✅ Live list closed.")
        await query.answer()
        return
    
    if data.startswith("live_copy_uid:"):
        uid = data.split(":")[1]
        await query.answer(f"✅ Copied!", show_alert=False)
        await query.message.reply_text(
            f"📋 **User ID Copied!**\n\n"
            f"<code>{uid}</code>\n\n"
            f"💡 Press and hold the code above to copy!",
            parse_mode="HTML"
        )
        return


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


async def cmd_rcvall(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send a single XLSX with all submissions from all users (no user info)."""
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    status_msg = await update.message.reply_text("⏳ Generating XLSX for all submissions...")

    tmp_path = None
    try:
        all_subs = get_all_submissions()
        if not all_subs:
            await status_msg.edit_text("📭 No submissions found.")
            return

        # Build XLSX in memory
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "All Submissions"

        # Header row — NO user info
        ws.append(["#", "Username", "Password", "2FA Key", "Date (UTC)"])

        row_num = 1
        for user_id, submissions in all_subs.items():
            if not isinstance(submissions, dict):
                continue
            for sub_id, sub in submissions.items():
                if not isinstance(sub, dict):
                    continue
                ws.append([
                    row_num,
                    sub.get("username", ""),
                    sub.get("password", ""),
                    sub.get("key", ""),
                    sub.get("datetime", ""),
                ])
                row_num += 1

        if row_num == 1:
            await status_msg.edit_text("📭 No submissions found.")
            return

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            wb.save(tmp.name)
            tmp_path = tmp.name

        with open(tmp_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename="all_submissions.xlsx",
                caption=f"📊 Total submissions: {row_num - 1}",
            )
        await status_msg.delete()

    except Exception as e:
        logger.error(f"cmd_rcvall failed: {e}")
        await status_msg.edit_text(f"❌ Error: {e}")
    finally:
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
    
async def callback_2fa_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle 2FA inline keyboard callbacks (copy only)."""
    query = update.callback_query
    data = query.data
    
    # Handle copy OTP (1-click copy) - shows code in a selectable message
    if data.startswith("copy_otp:"):
        otp_code = data.split(":")[1]
        await query.answer(f"✅ Here's your code!", show_alert=False)
        # Send code in a separate message that user can press and hold to copy
        await query.message.reply_text(
            f"📋 **Your 6-Digit Code:**\n"
            f"<code>{otp_code}</code>\n\n"
            f"💡 Press and hold on the code to copy it!",
            parse_mode="HTML"
        )
        return
    
    # Handle confirm registration
    if data == "confirm_registered":
        user = update.effective_user
        
        key = context.user_data.get("pending_2fa_key")
        ig_username = context.user_data.get("pending_username", generate_username())
        
        if not key:
            await query.answer("❌ Session expired!", show_alert=True)
            await query.edit_message_text("⚠️ Session expired. Please start the task again.")
            return
        
        password = "axiex@5"
        tg_username = f"@{user.username}" if user.username else str(user.id)
        
        try:
            add_submission(user.id, tg_username, ig_username, password, key)
        except Exception as e:
            logger.error(f"callback_2fa_handler add_submission failed: {e}")
            await query.answer("❌ Failed to save submission!", show_alert=True)
            return
        
        # Clear pending data
        context.user_data.pop("pending_2fa_key", None)
        context.user_data.pop("pending_username", None)
        context.user_data.pop("task_username", None)
        context.user_data.pop("pending_otp_code", None)
        
        await query.answer("✅ Account registered!")
        await query.edit_message_text(
            f"✅ **Account Successfully Registered!**\n\n"
            f"⏳ Review time: 6-12 hours\n\n"
            f"💰 Upon approval, you will receive `+${get_task_price():.4f}`\n\n"
            f"Thank you! 🎉",
            parse_mode=None
        )
        return
    
    # Handle cancel task
    if data == "cancel_2fa_task":
        context.user_data.pop("pending_2fa_key", None)
        context.user_data.pop("pending_username", None)
        context.user_data.pop("task_username", None)
        context.user_data.pop("pending_otp_code", None)
        
        await query.answer("❌ Task cancelled!")
        await query.edit_message_text(
            "❌ **Task Cancelled**\n\nYou can start a new task anytime.",
            parse_mode=None
        )
        return

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
        "📋 **ADMIN COMMANDS**\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n\n"
        
        "👥 **USER MANAGEMENT**\n"
        "/add {amount} {userid} — Balance barhao\n"
        "/rm {amount} {userid} — Balance kamo\n"
        "/apr {amount} {userid} — Approved count barhao\n"
        "/rmreview {userid} — Sob submissions delete\n"
        "/userinfo {userid} — User info dekho\n"
        "/list — Sob user list (copy)\n"
        "/acts {userid} — Pending submissions review\n\n"
        
        "📢 **MESSAGING**\n"
        "/msg {userid} {text} — Message pathao\n"
        "/cast {text} — Broadcast koro\n\n"
        
        "📊 **LEADERBOARD**\n"
        "/ldset — Real leaderboard (auto enables)\n"
        "/ldauto — Auto leaderboard (auto enables)\n"
        "/ldoff — Turn leaderboard OFF\n\n"
        
        "🤖 **BOT CONTROL**\n"
        "/botoff — Turn bot OFF (maintenance mode)\n"
        "/boton — Turn bot back ON\n\n"
        
        "📁 **SUBMISSIONS**\n"
        "/rcv {userid} — Excel file nao\n\n"
        
        "⚙️ **SETTINGS**\n"
        "/stp {price} — Task price set\n\n"
        
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "/cmd — Ei list"
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
            
            TASK_2FA_STARTED: [
    MessageHandler(filters.Regex("^✅ Account Registered$"), handle_account_registered),
    MessageHandler(filters.Regex("^❌ Cancel Task$"), handle_2fa_cancel),
            ],
    
            WITHDRAW_MENU: [
    MessageHandler(filters.Regex("^USDT-BEP20$"), handle_withdraw_bep20),
    MessageHandler(filters.Regex("^BKASH - BDT$"), handle_withdraw_bkash),  # NEW
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
    app.add_handler(CallbackQueryHandler(callback_live_handler, pattern="^(live_page:|live_copy_uid:|live_close)"))
    app.add_handler(CallbackQueryHandler(callback_check_join, pattern="^check_join$"))
    app.add_handler(CallbackQueryHandler(callback_acts_copy_handler, pattern="^(acts_copy_username:|acts_copy_password:|acts_copy_2fa:)"))
    app.add_handler(CallbackQueryHandler(callback_2fa_handler, pattern="^(copy_otp:|confirm_registered|cancel_2fa_task)"))
    app.add_handler(CallbackQueryHandler(callback_list_handlers, pattern="^(list_page:|copy_uid:|list_close)"))
    app.add_handler(CallbackQueryHandler(callback_acts_handler, pattern="^(acts_|acts_exit)"))
    app.add_handler(CommandHandler("live", cmd_live))
    app.add_handler(CommandHandler("rcv", cmd_rcv))
    app.add_handler(CommandHandler("rcvall", cmd_rcvall))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("rm", cmd_rm))
    app.add_handler(CommandHandler("msg", cmd_msg))
    app.add_handler(CommandHandler("ldoff", cmd_ldoff))
    app.add_handler(CommandHandler("cast", cmd_cast))
    app.add_handler(CommandHandler("userinfo", cmd_userinfo))
    app.add_handler(CommandHandler("stp", cmd_stp))
    app.add_handler(CommandHandler("ldset", cmd_ldset))
    app.add_handler(CommandHandler("ldauto", cmd_ldauto))
    app.add_handler(CommandHandler("acts", cmd_acts))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("botoff", cmd_botoff))
    app.add_handler(CommandHandler("boton", cmd_boton))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(CommandHandler("rmreview", cmd_rmreview))
    app.add_handler(CommandHandler("apr", cmd_apr))
    app.add_handler(CommandHandler("cmd", cmd_cmd))
    app.add_error_handler(error_handler)

    logger.info("Bot is starting (polling)…")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
