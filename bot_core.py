import os, time, re, json, base64, requests, msal, openai
from datetime import datetime

TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
OPENAI_KEY    = os.environ["OPENAI_KEY"]
MAIN_DIR      = "MainFolder"
CHECKLIST_DIR = f"{MAIN_DIR}/Checklist"
RESPONSE_DIR  = f"{MAIN_DIR}/EmailResponses"
ENGINE_KW     = ["engine","engine failure","engine damaged","engine fire","engine broken","engine rusted"]
openai.api_key = OPENAI_KEY

TOKEN = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET).acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"])["access_token"]
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}

def drive(path): return f"https://graph.microsoft.com/v1.0/me/drive/root:/{path}"
def ensure(path):
    if "error" in requests.get(drive(path), headers=HEADERS).json():
        requests.post(drive("").rstrip(":")+"/children", headers=HEADERS,
                      json={"name": path.split("/")[-1], "folder": {}})
def best_pdf(subject, body):
    files = requests.get(drive(f"{CHECKLIST_DIR}:/children"), headers=HEADERS).json().get("value",[])
    text  = (subject + " " + body).lower()
    for f in files:
        if f["name"].endswith(".pdf") and any(k in text for k in f["name"].lower().split()): return f["name"]
    return next((f["name"] for f in files if f["name"].endswith(".pdf")), None)
def gpt_reply(subject, body, sender):
    prompt = f"""You are a marine engineer. Subject: {subject} Body: {body} From: {sender}
Draft a professional reply that acknowledges the engine issue, guides through the attached checklist, offers further help. Sign: Tech Support Team"""
    return openai.ChatCompletion.create(model="gpt-3.5-turbo",
                                        messages=[{"role":"user","content":prompt}])["choices"][0]["message"]["content"]
def send_reply(msg, reply_body, pdf_name):
    sender, subj = msg["from"]["emailAddress"]["address"], msg["subject"]
    pdf_bytes = requests.get(drive(f"{CHECKLIST_DIR}/{pdf_name}")+":/content", headers=HEADERS).content
    payload = {"message": {"toRecipients": [{"emailAddress": {"address": sender}}],
                           "subject": f"Re: {subj}",
                           "body": {"contentType": "HTML", "content": reply_body.replace("\n", "<br>")},
                           "attachments": [
                               {"@odata.type": "#microsoft.graph.fileAttachment",
                                "name": pdf_name,
                                "contentBytes": base64.b64encode(pdf_bytes).decode(),
                                "contentType": "application/pdf"}]}}
    requests.post("https://graph.microsoft.com/v1.0/me/sendMail", headers=HEADERS, json=payload)
    print(f"[+] Sent to {sender}")
def save_thread(msg, reply, pdf):
    case_id = f"CASE-{int(time.time())}"
    ensure(f"{RESPONSE_DIR}/{case_id}")
    summary = {"caseId": case_id, "from": msg["from"]["emailAddress"]["address"],
               "subject": msg["subject"], "inbound": msg["body"]["content"],
               "outbound": reply, "checklist": pdf}
    requests.put(drive(f"{RESPONSE_DIR}/{case_id}/thread.json")+":/content",
                 headers=HEADERS, data=json.dumps(summary, indent=2))
def process_message(msg):
    subj, body = msg["subject"].lower(), msg.get("body",{}).get("content","").lower()
    if not any(k in subj+body for k in ENGINE_KW): return
    pdf_name = best_pdf(subj, body)
    if not pdf_name: print("[-] No checklist"); return
    reply = gpt_reply(msg["subject"], body, msg["from"]["emailAddress"]["address"])
    send_reply(msg, reply, pdf_name)
    save_thread(msg, reply, pdf_name)
    requests.patch(f"https://graph.microsoft.com/v1.0/me/messages/{msg['id']}", headers=HEADERS, json={"isRead": True})
def sync_mail():
    ensure(MAIN_DIR); ensure(CHECKLIST_DIR); ensure(RESPONSE_DIR)
    delta_url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta"
    while True:
        r = requests.get(delta_url, headers=HEADERS).json()
        for m in r.get("value", []):
            if not m.get("isRead", True): process_message(m)
        delta_url = r.get("@odata.nextLink") or r.get("@odata.deltaLink", delta_url)
        time.sleep(30)
