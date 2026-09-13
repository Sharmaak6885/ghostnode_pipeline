# GhostNode | Bulk Email Marketing & Lead Generation Pipeline

GhostNode is a full-stack, modern SaaS application designed for targeted B2B lead generation and bulk email outreach. With a sleek UI and powerful Python backend, GhostNode allows you to automatically scrape high-quality prospects from the web and launch personalized email campaigns seamlessly.

## Features

- 🎯 **Deep Web Lead Extraction:** Specify a target role (e.g., "CTO") and industry (e.g., "SaaS"). The built-in OSINT crawler utilizes DuckDuckGo and TLS fingerprint spoofing (`curl_cffi`) to bypass firewalls and extract verified emails and phone numbers.
- 🛡️ **Confidence Filtering:** Automatically ranks leads based on data completeness (e.g., presence of phone numbers) and lets you filter out low-confidence targets before outreach.
- 📧 **Automated Bulk Emailing:** Attach custom PDFs or CSV lists and deploy templated emails at scale using a dedicated SMTP daemon. 
- 🧪 **A/B Testing:** Input two different subject lines to randomize delivery and test which subject yields a higher open rate.
- 👀 **Live HTML Preview:** A WYSIWYG editor allows you to see exactly what your email will look like before you hit send.
- 📊 **Telemetry & Analytics:** Tracks real-time metrics including total emails sent, verified opens (via pixel tracking), bounces, and total opt-outs.
- 📖 **Campaign History Log:** Keeps a detailed, local database of every campaign launched for complete auditability.

## Folder Structure

```
ghostnode_pipeline/
│
├── app.py                  # Main Flask application and backend logic (Scraper, SMTP Engine)
├── templates/
│   └── index.html          # The single-page frontend Dashboard (HTML/CSS/JS)
├── static/
│   └── img/
│       └── logo.jpg        # Website brand logo
├── .env                    # (Not tracked) Stores SMTP credentials (SENDER_EMAIL, APP_PASSWORD)
├── .gitignore              # Defines confidential runtime files to keep off GitHub
└── README.md               # Project documentation
```

*(Note: During runtime, the app will generate local files such as `dynamic_leads.xlsx`, `campaign_stats.json`, `campaign_history.json`, `blacklist.txt`, and `settings.json` to store your local data safely without exposing it online).*

## Setup Instructions

### 1. Prerequisites
- **Python 3.8+** installed on your computer.
- A Gmail account with **App Passwords** enabled. (Standard passwords will not work for SMTP).

### 2. Clone the Repository
```bash
git clone https://github.com/Sharmaak6885/ghostnode_pipeline.git
cd ghostnode_pipeline
```

### 3. Install Dependencies
Install the required Python packages:
```bash
pip install flask python-dotenv curl_cffi duckduckgo_search openpyxl
```

### 4. Configure Credentials
Create a file named `.env` in the root of the project directory. Open it and add your Gmail SMTP credentials:
```env
SENDER_EMAIL="your_email@gmail.com"
APP_PASSWORD="your_16_digit_app_password"
```

### 5. Run the Application
Start the Flask server:
```bash
python app.py
```

### 6. Open the Dashboard
Open your web browser and navigate to:
```
http://127.0.0.1:5001
```

---
*Disclaimer: Ensure you comply with CAN-SPAM and local regulations when using automated bulk email and scraping tools.*