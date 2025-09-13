from fastapi import FastAPI
from threading import Thread
import bot_core

app = FastAPI(title="Outlook Engine Bot")
Thread(target=bot_core.sync_mail, daemon=True).start()

@app.get("/")
def health(): return {"status": "alive", "bot": "running"}
