import os
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

api_id_raw = os.environ.get("TELEGRAM_API_ID")
api_hash = os.environ.get("TELEGRAM_API_HASH")
session_str = os.environ.get("TELEGRAM_SESSION")

api_id = int(api_id_raw)

with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
    for dialog in client.iter_dialogs():
        if dialog.is_group or dialog.is_channel:
            print(f"{dialog.id}\t{dialog.name}")
