#!/usr/bin/env python3
"""
Instagram Account Buyer Bot
Production-ready Telegram bot for Railway + GitHub deployment.
"""

import asyncio
import json
import logging
import os
import random
import string
import tempfile
from datetime import datetime, timezone

import firebase_admin
import openpyxl
from firebase_admin import credentials, db
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.constants import ParseMode
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
) = range(8)

# ─────────────────────────────────────────────
# Firebase Initialisation
# ─────────────────────────────────────────────
def init_firebase():
    config = json.loads(FIREBASE_CONFIG)
    database_url = config.pop("databaseURL", None)
    if not database_url:
        raise ValueError("FIREBASE_CONFIG must contain 'databaseURL' key.")
    cred = credentials.Certificate(config)
    firebase_admin.initialize_app(cred, {"databaseURL": database_url})
    logger.info("Firebase initialised successfully.")

# ─────────────────────────────────────────────
# Firebase Helpers
# ─────────────────────────────────────────────
def get_user(user_id: int) -> dict:
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


def update_user(user_id: int, updates: dict):
    db.reference(f"users/{user_id}").update(updates)


def add_submission(user_id: int, tg_username: str, username: str, password: str, key: str):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    sub_ref = db.reference(f"submissions/{user_id}").push()
    sub_ref.set({
        "username": username,
        "password": password,
        "key": key,
        "tg_username": tg_username,
        "datetime": now,
    })
    # Update user counters
    user = get_user(user_id)
    update_user(user_id, {
        "in_review": user.get("in_review", 0) + 1,
        "total_submitted": user.get("total_submitted", 0) + 1,
    })


def get_submissions(user_id: int) -> list[dict]:
    data = db.reference(f"submissions/{user_id}").get()
    if not data:
        return []
    return [{"id": k, **v} for k, v in data.items()]


def remove_submissions(user_id: int):
    count = len(get_submissions(user_id))
    db.reference(f"submissions/{user_id}").delete()
    user = get_user(user_id)
    new_review = max(0, user.get("in_review", 0) - count)
    update_user(user_id, {"in_review": new_review})


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
    return ref.key  # withdrawal ID


def get_withdrawal(user_id: int, w_id: str) -> dict | None:
    return db.reference(f"withdrawals/{user_id}/{w_id}").get()


def update_withdrawal(user_id: int, w_id: str, updates: dict):
    db.reference(f"withdrawals/{user_id}/{w_id}").update(updates)


def get_all_users() -> dict:
    data = db.reference("users").get()
    return data or {}


def get_all_submissions() -> dict:
    data = db.reference("submissions").get()
    return data or {}

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


def generate_username() -> str:
    adj = random.choice(UNCOMMON_ADJECTIVES)
    noun = random.choice(UNCOMMON_NOUNS)
    num = random.randint(100, 9999)
    return f"{adj}_{noun}_{num}"


def generate_tx_hash() -> str:
    return "0x" + "".join(random.choices(string.hexdigits.lower(), k=64))


def build_xlsx(rows: list[dict]) -> bytes:
    """Build XLSX bytes from submission rows. Returns raw bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Submissions"
    ws.append(["Username", "Password", "2FA Key", "TG Username", "DateTime"])
    for row in rows:
        ws.append([
            row.get("username", ""),
            row.get("password", ""),
            row.get("key", ""),
            row.get("tg_username", ""),
            row.get("datetime", ""),
        ])
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb.save(tmp.name)
        tmp.seek(0)
        with open(tmp.name, "rb") as f:
            data = f.read()
    return data


# ─────────────────────────────────────────────
# Keyboards
# ─────────────────────────────────────────────
HOME_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["💰 Balance", "📋 Tasks"],
        ["📥 Withdraw", "🫟 I'm New User"],
        ["💝 Support"],
    ],
    resize_keyboard=True,
)

TASK_MENU_KEYBOARD = ReplyKeyboardMarkup(
    [["Inst 2FA - 0.024$"], ["🔙 Back"]],
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
    # Ensure user exists in Firebase
    get_user(user.id)
    full_name = user.full_name or user.first_name or "User"
    await update.message.reply_text(
        f"🥰 স্বাগতম, {full_name}!\n"
        "💎 কাজ শুরু করতে নিচের অপশনগুলো ব্যবহার করুন 🔽",
        reply_markup=HOME_KEYBOARD,
    )
    context.user_data.clear()
    return HOME

# ─────────────────────────────────────────────
# HOME handlers
# ─────────────────────────────────────────────
async def handle_balance(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_id = update.effective_user.id
    data = get_user(user_id)
    balance = data.get("balance", 0.0)
    approved = data.get("approved", 0)
    in_review = data.get("in_review", 0)

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


async def handle_new_user(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Coming Soon 🔜", reply_markup=HOME_KEYBOARD)
    return HOME


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
        f"👤 Username: {username}\n"
        f"🔓 Password: filesubmit@2\n\n"
        "📱 Open account with above username and password.\n"
        "Then submit account with 2FA Key below 😄",
        reply_markup=ReplyKeyboardMarkup([["Cancel ❌"]], resize_keyboard=True),
    )
    return TASK_2FA_AWAIT_KEY


async def handle_2fa_key(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    key = update.message.text.strip()
    user = update.effective_user

    # Cancel mid-way
    if key == "Cancel ❌":
        return await handle_task_cancel(update, context)

    if len(key) != 16:
        await update.message.reply_text(
            f"❌ Invalid 2FA key. Must be exactly 16 characters. You entered {len(key)}. Try again:",
        )
        return TASK_2FA_AWAIT_KEY

    ig_username = context.user_data.get("task_username", generate_username())
    password = "filesubmit@2"
    tg_username = f"@{user.username}" if user.username else str(user.id)

    add_submission(user.id, tg_username, ig_username, password, key)

    await update.message.reply_text(
        "Your Account was been Submitted ✅",
        reply_markup=HOME_KEYBOARD,
    )
    context.user_data.pop("task_username", None)
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
    data = get_user(user_id)
    balance = data.get("balance", 0.0)
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

    wallet = text
    amount = context.user_data.get("withdraw_amount", 0.0)
    fee = 0.025
    receive = round(amount - fee, 4)
    user = update.effective_user
    tg_username = f"@{user.username}" if user.username else str(user.id)

    w_id = create_withdrawal(user.id, tg_username, amount, wallet)
    context.user_data.clear()

    # Confirm to user
    await update.message.reply_text(
        f"✅ Withdrawal request created!\n\n"
        f"💳 Method: USDT (BEP-20)\n"
        f"👛 Wallet: {wallet}\n"
        f"💵 Debit amount: ${amount:.4f}\n"
        f"📉 Fee: $0.0250\n"
        f"💰 You will receive: ${receive:.4f}",
        reply_markup=HOME_KEYBOARD,
    )

    # Notify admins
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
    data = query.data  # wd_approve:{user_id}:{w_id} or wd_cancel:{...}

    try:
        action, user_id_str, w_id = data.split(":")
        user_id = int(user_id_str)
    except Exception:
        await query.edit_message_text("❌ Invalid callback data.")
        return

    wd = get_withdrawal(user_id, w_id)
    if not wd:
        await query.edit_message_text("❌ Withdrawal record not found.")
        return

    if wd.get("status") != "pending":
        await query.edit_message_text(
            f"ℹ️ This request is already {wd.get('status')}."
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
                text="❌ Withdrawal request cancelled.",
            )
        except Exception as e:
            logger.warning(f"Could not message user {user_id}: {e}")

    elif action == "wd_approve":
        amount = wd.get("amount", 0.0)
        receive = wd.get("receive", 0.0)
        wallet = wd.get("wallet", "")

        # Deduct balance
        user_data = get_user(user_id)
        new_balance = max(0.0, round(user_data.get("balance", 0.0) - amount, 4))
        update_user(user_id, {"balance": new_balance})
        update_withdrawal(user_id, w_id, {"status": "approved"})

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


async def cmd_live(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    all_users = get_all_users()
    all_subs = get_all_submissions()

    if not all_users:
        await update.message.reply_text("No users found.")
        return

    lines = ["📊 Live User Stats\n━━━━━━━━━━━━━━━━━━━━━━"]
    for uid, udata in all_users.items():
        sub_count = len(all_subs.get(uid, {}))
        lines.append(
            f"🆔 {uid} | Submitted: {sub_count} | Balance: ${udata.get('balance', 0.0):.4f}"
        )

    await update.message.reply_text("\n".join(lines))


async def cmd_rcv(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("❌ Unauthorised.")
        return

    args = context.args
    if not args or len(args) < 1:
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

    xlsx_bytes = build_xlsx(subs)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(xlsx_bytes)
        tmp.flush()
        tmp.seek(0)
        with open(tmp.name, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"submissions_{target_id}.xlsx",
                caption=f"📁 Submissions for user {target_id} ({len(subs)} records)",
            )


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

    user_data = get_user(target_id)
    new_balance = round(user_data.get("balance", 0.0) + amount, 4)
    update_user(target_id, {"balance": new_balance})
    await update.message.reply_text(
        f"✅ Added ${amount:.4f} to user {target_id}.\nNew balance: ${new_balance:.4f}"
    )


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

    user_data = get_user(target_id)
    new_balance = max(0.0, round(user_data.get("balance", 0.0) - amount, 4))
    update_user(target_id, {"balance": new_balance})
    await update.message.reply_text(
        f"✅ Removed ${amount:.4f} from user {target_id}.\nNew balance: ${new_balance:.4f}"
    )


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

    remove_submissions(target_id)
    update_user(target_id, {"in_review": 0})
    await update.message.reply_text(
        f"✅ All submissions removed for user {target_id}."
    )


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

    user_data = get_user(target_id)
    new_approved = user_data.get("approved", 0) + amount
    # Also reduce in_review by same amount (capped at 0)
    new_review = max(0, user_data.get("in_review", 0) - amount)
    update_user(target_id, {"approved": new_approved, "in_review": new_review})
    await update.message.reply_text(
        f"✅ Added {amount} to approved count for user {target_id}.\n"
        f"Total approved: {new_approved}"
    )

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
# Main
# ─────────────────────────────────────────────
def main():
    init_firebase()

    app = Application.builder().token(BOT_TOKEN).build()

    # ── Conversation Handler ──
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            HOME: [
                MessageHandler(filters.Regex("^💰 Balance$"), handle_balance),
                MessageHandler(filters.Regex("^📋 Tasks$"), handle_tasks),
                MessageHandler(filters.Regex("^📥 Withdraw$"), handle_withdraw),
                MessageHandler(filters.Regex("^🫟 I'm New User$"), handle_new_user),
                MessageHandler(filters.Regex("^💝 Support$"), handle_support),
            ],
            TASK_MENU: [
                MessageHandler(filters.Regex("^Inst 2FA - 0.024\\$$"), handle_task_2fa_info),
                MessageHandler(filters.Regex("^🔙 Back$"), handle_back_to_home),
            ],
            TASK_2FA_INFO: [
                MessageHandler(filters.Regex("^✨ Start$"), handle_task_start),
                MessageHandler(filters.Regex("^Cancel ❌$"), handle_task_cancel),
                MessageHandler(filters.Regex("^🔙 Back$"), handle_back_to_home),
            ],
            TASK_2FA_AWAIT_KEY: [
                MessageHandler(filters.Regex("^Cancel ❌$"), handle_task_cancel),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_2fa_key),
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

    # ── Inline Callbacks ──
    app.add_handler(CallbackQueryHandler(callback_withdrawal, pattern="^wd_(approve|cancel):"))

    # ── Admin Commands ──
    app.add_handler(CommandHandler("live", cmd_live))
    app.add_handler(CommandHandler("rcv", cmd_rcv))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("rm", cmd_rm))
    app.add_handler(CommandHandler("rmreview", cmd_rmreview))
    app.add_handler(CommandHandler("apr", cmd_apr))

    # ── Error Handler ──
    app.add_error_handler(error_handler)

    logger.info("Bot is starting (polling)…")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
