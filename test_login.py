from telethon import TelegramClient
import asyncio

api_id = 35675172
api_hash = "8e8b2d999e3359a003983dd57c8e771b"

session_name = "my_telegram_session"

async def main():
    client = TelegramClient(session_name, api_id, api_hash)
    await client.start()

    print("✅ 登录成功")
    print("正在读取聊天列表...\n")

    dialogs = await client.get_dialogs()

    for i, dialog in enumerate(dialogs, start=1):
        print(f"{i}. {dialog.name}")

    await client.disconnect()

if __name__ == "__main__":
    asyncio.run(main())