# VRCOMM LINE Bot

A webhook server that connects your **LINE Official Account** to Claude AI.

**What it does:**
- Receives every LINE message sent to your OA
- Logs to **SQLite** (local, fast) AND **Google Sheets** (persistent, always accessible)
- Processes the message with Claude AI (VRCOMM Admin persona)
- Sends an AI-generated reply back to the customer
- Exports all logged messages as a downloadable Excel file

---

## File Structure

```
vrcomm-line-bot/
├── app.py            Main Flask webhook server
├── db.py             SQLite message logger
├── ai_handler.py     Claude AI reply generator
├── excel_export.py   Excel (.xlsx) export engine
├── sheets_logger.py  Google Sheets real-time logger
├── requirements.txt  Python dependencies
├── render.yaml       Render.com deployment config
└── README.md         This file
```

---

## Step 1 — Get your LINE credentials

1. Go to [LINE Developers Console](https://developers.line.biz/console/)
2. Select your channel → **Messaging API** tab
3. Note down:
   - **Channel Secret** (Basic settings tab)
   - **Channel Access Token** (Messaging API tab → Issue if not already)

---

## Step 2 — Get your Anthropic API Key

1. Go to [console.anthropic.com](https://console.anthropic.com/)
2. Create an API key under **API Keys**

---

## Step 3 — Set up Google Sheets logging (recommended)

This keeps all your LINE messages permanently — even when Render restarts.

### 3a. Create a Google Cloud Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use existing) — name it e.g. `vrcomm-linebot`
3. Go to **APIs & Services → Enable APIs**
   - Enable **Google Sheets API**
   - Enable **Google Drive API**
4. Go to **APIs & Services → Credentials → Create Credentials → Service Account**
   - Name: `vrcomm-linebot-sa`
   - Role: **Editor**
   - Click Done
5. Click the service account → **Keys tab → Add Key → Create new key → JSON**
   - A `.json` file downloads to your computer — keep it safe

### 3b. Create the Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) → Create a new blank spreadsheet
2. Name it: `VRCOMM LINE Messages`
3. Copy the **Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SHEET_ID_IS_HERE/edit
   ```

### 3c. Share the Sheet with the service account

1. Open the `.json` key file you downloaded — find the `client_email` field, e.g.:
   ```
   vrcomm-linebot-sa@vrcomm-linebot.iam.gserviceaccount.com
   ```
2. In Google Sheets → click **Share**
3. Paste that email address → set role to **Editor** → click Send

### 3d. Prepare the credentials for Render

Open the `.json` key file in Notepad — it looks like:
```json
{
  "type": "service_account",
  "project_id": "vrcomm-linebot",
  "private_key_id": "...",
  ...
}
```
You'll paste this entire JSON (all on one line or as-is) into a Render environment variable.

---

## Step 4 — Deploy to Render.com

### 4a. Push code to GitHub

```bash
cd vrcomm-line-bot
git init
git add .
git commit -m "VRCOMM LINE Bot initial commit"
# Create a new repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/vrcomm-line-bot.git
git push -u origin main
```

### 4b. Create Web Service on Render

1. Go to [render.com](https://render.com) → **New → Web Service**
2. Connect your GitHub repo
3. Render auto-detects `render.yaml` — confirm settings
4. **Start Command:**
   ```
   gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 60
   ```
5. Set these **Environment Variables** in the Render dashboard:

| Key | Value |
|-----|-------|
| `LINE_CHANNEL_SECRET` | Your LINE Channel Secret |
| `LINE_CHANNEL_ACCESS_TOKEN` | Your LINE Channel Access Token |
| `ANTHROPIC_API_KEY` | Your Anthropic API Key |
| `GOOGLE_CREDENTIALS_JSON` | Paste the full contents of the service account `.json` file |
| `GOOGLE_SHEET_ID` | The Sheet ID from the spreadsheet URL |
| `EXPORT_PASSWORD` | (Optional) Any password to protect /export |

6. Click **Deploy**
7. Wait ~2 min — your URL will be: `https://vrcomm-line-bot.onrender.com`

---

## Step 5 — Set Webhook URL in LINE Console

1. Go to LINE Developers Console → your channel → **Messaging API** tab
2. Set **Webhook URL** to:
   ```
   https://vrcomm-line-bot.onrender.com/webhook
   ```
3. Enable **"Use webhook"** toggle → ON
4. Click **Verify** — should show ✅ Success
5. (Optional) Disable **"Auto-reply messages"** so only AI replies are sent

---

## Step 6 — Test it

1. Open LINE on your phone
2. Add your LINE OA as a friend (scan QR from LINE Console)
3. Send a message like `"สอบถามราคา Fortinet ครับ"`
4. The bot replies in ~3 seconds
5. Open your Google Sheet — the message should appear instantly

---

## Viewing your data

### Google Sheets (live, persistent)
Your Google Sheet auto-updates with every message. Columns logged:

| Column | Data |
|---|---|
| No. | Row number |
| Timestamp (LINE) | When LINE delivered the message |
| Logged At (Server) | When our server received it |
| Display Name | Customer's LINE name |
| User ID | Permanent LINE user ID |
| Source Type | user / group / room |
| Source ID | Group or room ID |
| Message Type | text / image / sticker / audio / video |
| Message | Message text content |
| Detail | Full detail (IDs, dimensions for media) |
| Message ID | LINE message ID |
| Reply Token | One-time reply token |
| AI Reply | What the bot replied back |

### Excel Download (snapshot)
```
https://vrcomm-line-bot.onrender.com/export
```
With password:
```
https://vrcomm-line-bot.onrender.com/export?password=YOUR_PASSWORD
```

### Health Check
```
https://vrcomm-line-bot.onrender.com/
```
Returns: service status + total messages logged in current session.

---

## Important Notes on Render Free Tier

- **Spins down after 15 min of inactivity** — first message after idle takes ~30 sec to respond
- **SQLite resets on restart** — but Google Sheets data is permanent ✅
- To avoid cold-start delays, use [UptimeRobot](https://uptimerobot.com) to ping your URL every 10 min for free
