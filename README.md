# Instagram Account Buyer Bot 🤖

A production-ready Telegram bot built with **python-telegram-bot v21+**, **Firebase Realtime Database**, hosted on **Railway**.

---

## Setup

### 1. Clone & push to GitHub

```bash
git init
git add .
git commit -m "initial commit"
git remote add origin [github.com](https://github.com/YOUR_USERNAME/YOUR_REPO.git)
git push -u origin main
```

### 2. Firebase Setup

1. Go to [Firebase Console](https://console.firebase.google.com/)
2. Create a project → Enable **Realtime Database**
3. Go to **Project Settings → Service Accounts → Generate new private key**
4. Download the JSON file
5. Open it and **add** a `databaseURL` key:

```json
{
  "type": "service_account",
  "project_id": "...",
  "databaseURL": "[your-project-default-rtdb.firebaseio.com](https://YOUR-PROJECT-default-rtdb.firebaseio.com)",
  ...rest of keys...
}
```

6. Minify the entire JSON to a single line — this becomes your `FIREBASE_CONFIG` env variable.

### 3. Railway Deployment

1. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
2. Select your repo
3. Add environment variables:

| Variable          | Value                        |
|-------------------|------------------------------|
| `BOT_TOKEN`       | Your Telegram Bot Token      |
| `FIREBASE_CONFIG` | The minified JSON string     |

4. Railway auto-detects `Procfile` and runs `python main.py`.

---

## Admin Commands

| Command                  | Description                          |
|--------------------------|--------------------------------------|
| `/live`                  | Show all users + submission counts   |
| `/rcv {userid}`          | Get user's XLSX submission file      |
| `/add {amount} {userid}` | Add balance to a user                |
| `/rm {amount} {userid}`  | Remove balance from a user           |
| `/rmreview {userid}`     | Clear all pending submissions        |
| `/apr {amount} {userid}` | Add to approved count                |

---

## Firebase Structure

```
users/
  {user_id}/
    balance: 0.0
    approved: 0
    in_review: 0
    total_submitted: 0

submissions/
  {user_id}/
    {push_id}/
      username: ""
      password: "filesubmit@2"
      key: ""
      tg_username: ""
      datetime: ""

withdrawals/
  {user_id}/
    {push_id}/
      tg_username: ""
      amount: 0.0
      fee: 0.025
      receive: 0.0
      wallet: ""
      status: "pending|approved|cancelled"
      datetime: ""
```

