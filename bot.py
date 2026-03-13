import re
import json
import os

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, filters

from config import BOT_TOKEN
from tg_report_reader import (
    get_today_realtime_text,
    get_today_summary_text,
    get_group_check_text,
    get_single_group_text,
    export_today_excel,
    export_yesterday_excel,
    get_system_status_text,
)

# =========================
# 基础设定
# =========================
OWNERS = [
    "rogerben717",
    "mlys94"
]
ADMINS_FILE = "admins.json"

MENU = [
    ["📊 当前实时数据", "📈 今日汇总"],
    ["📤 今日导出Excel", "📤 昨日导出Excel"],
    ["🔍 查询单个BCSG", "📡 检查所有BC群"],
    ["⚙️ 系统状态"]
]

WAITING_FOR_BCSG = {}


# =========================
# 管理员资料读写
# =========================
def load_admins():
    if not os.path.exists(ADMINS_FILE):
        return set()

    try:
        with open(ADMINS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        if isinstance(data, list):
            return set(x.lower() for x in data if isinstance(x, str))

        return set()
    except Exception:
        return set()


def save_admins(admins):
    with open(ADMINS_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(list(admins)), f, ensure_ascii=False, indent=2)


ADMINS = load_admins()


# =========================
# 权限判断
# =========================
def get_username(update: Update):
    user = update.effective_user
    if not user or not user.username:
        return None
    return user.username.lower()


def is_owner(update: Update):
    username = get_username(update)
    return username in [x.lower() for x in OWNERS]


def is_admin(update: Update):
    username = get_username(update)
    if not username:
        return False
    return username in [x.lower() for x in OWNERS] or username in ADMINS


async def deny_access(update: Update):
    await update.message.reply_text("❌ 无权限使用此机器人")


# =========================
# 开始
# =========================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await deny_access(update)
        return

    keyboard = ReplyKeyboardMarkup(MENU, resize_keyboard=True)
    await update.message.reply_text(
        "🤖 BC Auto Report\n\n请选择功能：",
        reply_markup=keyboard
    )


# =========================
# Owner 管理员命令
# =========================
async def add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global ADMINS

    if not is_owner(update):
        await update.message.reply_text("❌ 只有 Owner 可以添加管理员")
        return

    if not context.args:
        await update.message.reply_text("用法：/addadmin @username")
        return

    username = context.args[0].strip().replace("@", "").lower()

    if not username:
        await update.message.reply_text("请输入正确 username，例如：/addadmin @Whoglobal")
        return

    # 修正：如果输入的是 owner，不需要再加成 admin
    if username in [o.lower() for o in OWNERS]:
        await update.message.reply_text("Owner 不需要添加为管理员。")
        return

    ADMINS.add(username)
    save_admins(ADMINS)

    await update.message.reply_text(f"✅ 已添加管理员：@{username}")


async def del_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global ADMINS

    if not is_owner(update):
        await update.message.reply_text("❌ 只有 Owner 可以删除管理员")
        return

    if not context.args:
        await update.message.reply_text("用法：/deladmin @username")
        return

    username = context.args[0].strip().replace("@", "").lower()

    if not username:
        await update.message.reply_text("请输入正确 username，例如：/deladmin @Whoglobal")
        return

    if username in ADMINS:
        ADMINS.remove(username)
        save_admins(ADMINS)
        await update.message.reply_text(f"✅ 已删除管理员：@{username}")
    else:
        await update.message.reply_text(f"⚠️ @{username} 不是管理员")


async def list_admins(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("❌ 只有 Owner 可以查看管理员")
        return

    lines = []
    lines.append("👑 Owners")
    for o in OWNERS:
        lines.append(f"@{o}")
    lines.append("")

    lines.append("👤 Admins")
    if ADMINS:
        for username in sorted(ADMINS):
            lines.append(f"@{username}")
    else:
        lines.append("暂无管理员")

    await update.message.reply_text("\n".join(lines))


# =========================
# 普通消息处理
# =========================
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await deny_access(update)
        return

    user_id = update.effective_user.id
    text = (update.message.text or "").strip()

    # 等待输入单个BCSG编号
    if WAITING_FOR_BCSG.get(user_id):
        code = text.upper().replace(" ", "")

        if not re.fullmatch(r"BCSG\d+", code):
            await update.message.reply_text("请输入正确格式，例如：BCSG60")
            return

        wait_msg = await update.message.reply_text(f"正在获取 {code} 数据，请稍等...")
        try:
            result = await get_single_group_text(code)
            await wait_msg.edit_text(result)
        except Exception as e:
            await wait_msg.edit_text(f"查询失败：{str(e)}")
        finally:
            WAITING_FOR_BCSG.pop(user_id, None)
        return

    if text == "📊 当前实时数据":
        wait_msg = await update.message.reply_text("正在获取数据，请稍等...")
        try:
            result = await get_today_realtime_text()
            await wait_msg.edit_text(result)
        except Exception as e:
            await wait_msg.edit_text(f"获取失败：{str(e)}")

    elif text == "📈 今日汇总":
        wait_msg = await update.message.reply_text("正在获取数据，请稍等...")
        try:
            result = await get_today_summary_text()
            await wait_msg.edit_text(result)
        except Exception as e:
            await wait_msg.edit_text(f"获取失败：{str(e)}")

    elif text == "📡 检查所有BC群":
        wait_msg = await update.message.reply_text("正在获取数据，请稍等...")
        try:
            result = await get_group_check_text()
            await wait_msg.edit_text(result)
        except Exception as e:
            await wait_msg.edit_text(f"检查失败：{str(e)}")

    elif text == "🔍 查询单个BCSG":
        WAITING_FOR_BCSG[user_id] = True
        await update.message.reply_text("请输入群编号，例如：BCSG60")

    elif text == "📤 今日导出Excel":
        wait_msg = await update.message.reply_text("正在生成今日 Excel，请稍等...")
        try:
            filename = await export_today_excel()
            with open(filename, "rb") as f:
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=f
                )
            await wait_msg.edit_text("✅ 今日 Excel 已生成")
        except Exception as e:
            await wait_msg.edit_text(f"导出失败：{str(e)}")

    elif text == "📤 昨日导出Excel":
        wait_msg = await update.message.reply_text("正在生成昨日 Excel，请稍等...")
        try:
            filename = await export_yesterday_excel()
            with open(filename, "rb") as f:
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=f
                )
            await wait_msg.edit_text("✅ 昨日 Excel 已生成")
        except Exception as e:
            await wait_msg.edit_text(f"导出失败：{str(e)}")

    elif text == "⚙️ 系统状态":
        wait_msg = await update.message.reply_text("正在检查系统状态，请稍等...")
        try:
            result = await get_system_status_text()
            await wait_msg.edit_text(result)
        except Exception as e:
            await wait_msg.edit_text(f"状态检查失败：{str(e)}")

    else:
        await update.message.reply_text("请使用菜单按钮操作。")


# =========================
# 主程序
# =========================
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("addadmin", add_admin))
    app.add_handler(CommandHandler("deladmin", del_admin))
    app.add_handler(CommandHandler("admins", list_admins))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("✅ BC Auto Report 已启动")
    app.run_polling()


if __name__ == "__main__":
    main()
