# 📧 Gmail Email Extractor

A powerful LangGraph-based multi-agent system that extracts all your sent emails from Gmail and exports them to an Excel file.

## Features

✨ **Multi-Agent Architecture** - Three specialized agents work together to extract and process your emails
- **Agent 1: Gmail Extractor** - Connects to Gmail and fetches all sent emails
- **Agent 2: Email Parser** - Extracts email addresses, subjects, dates, and other details
- **Agent 3: Excel Maker** - Creates a beautifully formatted Excel file with all the data

🚀 **Easy Web Interface** - Simple, clean UI to start the workflow
📊 **Clean Email Data** - Automatically extracts clean email addresses from messy headers
💾 **Excel Export** - Downloads ready-to-use Excel file with formatting

## Getting Your Gmail App Password

Gmail blocks direct password access for security. You need to generate an **App Password**:

### Step 1: Enable 2-Step Verification
1. Go to [myaccount.google.com](https://myaccount.google.com)
2. Click **Security** in the left sidebar
3. Find **2-Step Verification** and enable it (if not already enabled)

### Step 2: Generate App Password
1. Go back to **Security** settings
2. Find **App passwords** (only shows if 2-Step is enabled)
3. Select:
   - App: **Mail**
   - Device: **Windows Computer** (or your device)
4. Google will generate a **16-character password**
5. Copy this password and paste it in the web app

> ⚠️ **Note:** Use this app password, NOT your regular Gmail password

## Installation

### Requirements
- Python 3.8+
- Flask
- openpyxl
- langgraph
- imaplib (built-in)

### Setup

1. **Clone the repository**
```bash
git clone <repo-url>
cd gmail-extractor
```

2. **Create virtual environment**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install flask openpyxl langgraph
```

4. **Run the application**
```bash
python web_app.py
```

5. **Open browser**
```
http://localhost:5000
```

## How It Works

The system uses a **3-agent workflow** to extract and process your emails:

### 🔌 Agent 1: Gmail Extractor
- Connects to your Gmail account via IMAP
- Fetches all emails from your "Sent Mail" folder
- Returns a list of email IDs for processing

### 📝 Agent 2: Email Parser
- Takes the email IDs from Agent 1
- Extracts important details:
  - **Recipient email address** (cleaned from display names)
  - **Subject line**
  - **Date sent**
  - **Email ID**
- Handles messy email headers and multiple recipients

### 📑 Agent 3: Excel Maker
- Creates a professional Excel workbook
- Formats headers with blue background and white text
- Adjusts column widths for readability
- Saves file to Desktop as `sent_emails.xlsx`

## Usage

1. Enter your **Gmail account** and **App Password**
2. Click **Start Workflow**
3. Watch the live logs as agents process your emails
4. Download the Excel file when complete

## File Structure

```
├── web_app.py          # Flask server
├── app.py              # LangGraph workflow & agents
├── templates/
│   └── index.html      # Web interface
└── README.md           # This file
```

## Example Output

The Excel file will contain:

| Email ID | To | Subject | Date |
|----------|----|---------| ---- |
| 12345 | user@example.com | Hello | Thu, 25 Dec 2025 |
| 12346 | another@domain.com | Meeting Notes | Thu, 25 Dec 2025 |

## Troubleshooting

**"Connection failed" error**
- Verify your Gmail app password is correct
- Make sure 2-Step Verification is enabled
- Check that IMAP is enabled in Gmail settings

**"No emails to parse"**
- Make sure you have sent emails in your Gmail account
- The app limits to the last 10 emails to avoid processing delays

**Excel file not downloading**
- Check if the file was created on your Desktop
- Clear browser cache and try again

