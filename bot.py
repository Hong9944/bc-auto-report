import re
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

MENU = [
    ["📊 当前实时数据", "📈 今日汇总"],
    ["📤 今日导出Excel", "📤 昨日导出Excel"],
    ["🔍 查询单个BCSG", "📡 检查所有BC群"],
    ["⚙️ 系统状态"]
]

WAITING_FOR_BCSG = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = ReplyKeyboardMarkup(MENU, resize_keyboard=True)
    await update.message.reply_text(
        "🤖 BC Auto Report\n\n请选择功能：",
        reply_markup=keyboard
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
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


def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("✅ BC Auto Report 已启动")
    app.run_polling()


if __name__ == "__main__":
    main()
