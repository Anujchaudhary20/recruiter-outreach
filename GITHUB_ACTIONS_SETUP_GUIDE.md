# GitHub Actions Setup Guide - Automated Recruiter Outreach

## Overview

- **Runs**: Daily at 10:30 AM IST (05:00 AM UTC)
- **Fully Automated**: No manual intervention needed
- **Cost**: FREE
- **Action**: Scrapes recruiters → Tailors resumes → Sends emails → Updates Excel

---

## SETUP STEPS (30 minutes total)

### Step 1: Create GitHub Repository (5 minutes)

1. Go to **github.com** and login
2. Click **"+" icon** → **"New repository"**
3. Name: `recruiter-outreach`
4. Description: "Daily automated LinkedIn recruiter outreach"
5. Choose **Public** or **Private**
6. Click **"Create repository"**

---

### Step 2: Add Files to Repository (5 minutes)

Clone the repo locally:
```bash
git clone https://github.com/YOUR_USERNAME/recruiter-outreach.git
cd recruiter-outreach
```

Add these files to the repo:

1. **automated_outreach.py** (main script)
2. **requirements.txt** (Python dependencies)
3. **.github/workflows/outreach.yml** (GitHub Actions workflow)
4. **resume_template.txt** (your resume)
5. **recruiter_tracking.xlsx** (Excel tracker)

Create folder structure:
```
recruiter-outreach/
├── .github/
│   └── workflows/
│       └── outreach.yml
├── automated_outreach.py
├── requirements.txt
├── resume_template.txt
└── recruiter_tracking.xlsx
```

Push to GitHub:
```bash
git add .
git commit -m "Initial commit: automated outreach system"
git push origin main
```

---

### Step 3: Get Gmail API Credentials (10 minutes)

#### 3a. Create Google Cloud Project

1. Go to **https://console.cloud.google.com**
2. Click **"Create Project"**
3. Name: `Recruiter Outreach`
4. Click **"Create"**

#### 3b. Enable Gmail API

1. In left sidebar, go to **"APIs & Services"** → **"Library"**
2. Search for **"Gmail API"**
3. Click **Gmail API**
4. Click **"Enable"**

#### 3c. Create Service Account (Alternative: OAuth)

**For OAuth (Recommended for email sending):**

1. Go to **APIs & Services** → **Credentials**
2. Click **"+ Create Credentials"** → **"OAuth 2.0 Client IDs"**
3. Choose **"Desktop application"**
4. Download JSON file
5. Copy entire JSON content

**Save the JSON content** - you'll need it for Step 4

---

### Step 4: Add GitHub Secrets (5 minutes)

1. Go to your GitHub repo
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **"New repository secret"**

Add these secrets:

#### Secret 1: FIRECRAWL_API_KEY
- **Name**: `FIRECRAWL_API_KEY`
- **Value**: Your Firecrawl API key (from Claude connector)
- Click **"Add secret"**

#### Secret 2: GMAIL_CREDENTIALS_JSON
- **Name**: `GMAIL_CREDENTIALS_JSON`
- **Value**: Paste entire JSON content from Step 3c
- Click **"Add secret"**

#### Secret 3: ANUJ_EMAIL
- **Name**: `ANUJ_EMAIL`
- **Value**: `Anuj04004@gmail.com`
- Click **"Add secret"**

#### Secret 4: ANUJ_PHONE
- **Name**: `ANUJ_PHONE`
- **Value**: `+91-9528034629`
- Click **"Add secret"**

#### Secret 5 (Optional): SLACK_WEBHOOK
- **Name**: `SLACK_WEBHOOK`
- **Value**: Slack webhook URL (if you want notifications)
- Click **"Add secret"**

---

### Step 5: Enable GitHub Actions (2 minutes)

1. Go to repo **Settings** → **Actions** → **General**
2. Under **"Actions permissions"**, select **"Allow all actions and reusable workflows"**
3. Click **"Save"**

---

### Step 6: Test Workflow (3 minutes)

1. Go to **Actions** tab
2. Click **"Daily LinkedIn Recruiter Outreach"** workflow
3. Click **"Run workflow"** → **"Run workflow"** (manual trigger)
4. Wait for it to run (~2 minutes)
5. Check logs to confirm success

---

## SCHEDULE EXPLAINED

### Current Schedule: Daily at 10:30 AM IST

```
Cron: 0 5 * * *
      ↓ ↓ ↓ ↓ ↓
      │ │ │ │ └── Day of week (0-6, 0=Sunday)
      │ │ │ └──── Month (1-12)
      │ │ └────── Day (1-31)
      │ └──────── Hour (0-23)
      └────────── Minute (0-59)

0 5 * * * = Every day at 05:00 UTC (10:30 AM IST)
```

### Change Schedule (Optional):

Edit `.github/workflows/outreach.yml`:

```yaml
on:
  schedule:
    - cron: '0 5 * * *'  # Change this line
```

Examples:
- **12:00 PM IST**: `30 6 * * *`
- **2:00 PM IST**: `30 8 * * *`
- **4:00 PM IST**: `30 10 * * *`

---

## WHAT HAPPENS DAILY (10:30 AM IST)

### Automatic Execution:

1. **10:30 AM**: GitHub Actions trigger
2. **10:35 AM**: Firecrawl scrapes LinkedIn
3. **10:40 AM**: Extract 40-50 recruiters (India only)
4. **10:45 AM**: Tailor 40-50 resumes
5. **10:50 AM**: Generate personalized emails
6. **11:00 AM**: Send all via Gmail
7. **11:15 AM**: Update Excel tracker
8. **11:20 AM**: Log saved in Actions

**Your effort**: 0 minutes

---

## MONITORING

### Check Execution:

1. Go to **Actions** tab
2. Click latest workflow run
3. See logs + status
4. Download Excel tracker (artifacts)

### View Logs:

Each run shows:
- ✅ Recruiters found
- ✅ Emails sent
- ❌ Failures (if any)
- ⏰ Execution time

---

## TROUBLESHOOTING

### Issue: "Gmail authentication failed"

**Solution:**
1. Regenerate Gmail credentials (Step 3c)
2. Update `GMAIL_CREDENTIALS_JSON` secret
3. Re-run workflow

### Issue: "Firecrawl API key invalid"

**Solution:**
1. Get new API key from Claude connector
2. Update `FIRECRAWL_API_KEY` secret
3. Re-run workflow

### Issue: "Workflow not running at scheduled time"

**Solution:**
1. GitHub Actions requires repo to have had activity in past 60 days
2. Make a commit or manual workflow run
3. Scheduling should resume

### Issue: "No recruiters found"

**Solution:**
1. Check Firecrawl API response
2. Verify LinkedIn search queries
3. Try manual trigger to debug

---

## MANUAL TRIGGER (Anytime)

Want to run immediately without waiting for schedule?

1. Go to **Actions** tab
2. Click **"Daily LinkedIn Recruiter Outreach"**
3. Click **"Run workflow"** dropdown
4. Click **"Run workflow"** button

Runs within 1-2 minutes ✅

---

## DAILY OUTPUT

### In GitHub Actions (Actions tab):
- ✅ Execution logs
- ✅ Success/failure status
- ✅ Excel tracker (downloadable)

### In Gmail:
- ✅ "Sent" folder shows all emails (40-50 per day)
- ✅ Organized by date

### In Excel:
- ✅ recruiter_tracking.xlsx updated
- ✅ New rows for each day
- ✅ Download from GitHub Actions artifacts

---

## 30-DAY PROJECTION

| Metric | Value |
|--------|-------|
| Recruiters per day | 40-50 |
| Total in 30 days | 1,200-1,500 |
| All from India | ✅ |
| All with tailored resumes | ✅ |
| All sent via Gmail | ✅ |
| Cost | $0 |
| Your manual effort | 0 minutes |

---

## SUMMARY

✅ **Setup**: 30 minutes (one-time)
✅ **Maintenance**: 0 minutes/day
✅ **Automation**: Fully scheduled
✅ **Cost**: $0
✅ **Results**: 1,200-1,500 outreach in 30 days

---

## NEXT STEPS

1. Create GitHub repo (Step 1)
2. Add files (Step 2)
3. Get Gmail credentials (Step 3)
4. Add secrets (Step 4)
5. Enable actions (Step 5)
6. Test (Step 6)

**🎉 Done! System runs automatically starting tomorrow at 10:30 AM IST**

---

## QUESTIONS?

If workflow fails:
1. Check **Actions** logs
2. Verify secrets are correct
3. Ensure Gmail credentials are fresh
4. Retry with manual trigger

GitHub Actions has excellent documentation: **https://docs.github.com/en/actions**

👊 You're all set!
