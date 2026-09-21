# 🔗 TPB Broken Link Checker

> Automated monthly crawl of https://www.tpb.gov.au checking for broken links,
> server errors, and redirect issues. Emails a formatted Excel report automatically.

---

## 📋 What It Checks

| Domain | Included? |
|--------|-----------|
| `www.tpb.gov.au` | ✅ Crawled (all pages) |
| External links found on TPB pages | ✅ Checked (not crawled) |

| Error Type | Flagged? |
|------------|----------|
| 🔴 404 Not Found | ✅ Yes |
| 🔴 Other 4xx Client Errors | ✅ Yes |
| 🟠 5xx Server Errors | ✅ Yes |
| 🟡 Redirect Chains & Loops | ✅ Yes |
| ⏱️ Timeouts & Connection Errors | ✅ Yes |

---

## 🖱️ How to Run (No Tech Knowledge Needed)

### Option A — Automatic Monthly Run
Nothing to do! Runs automatically on the **28th of every month at 8am AEST**. ✅

### Option B — Run Manually (on demand)

1. Go to your GitHub repository
2. Click the **"Actions"** tab
3. Click **"🔗 TPB Monthly Link Checker"** in the left panel
4. Click **"Run workflow"** → green **"Run workflow"** button
5. ☕ Wait ~30–60 minutes for the crawl to finish
6. Check your **email** for the report

### Option C — Download from GitHub (if email fails)

1. Go to **Actions** tab → click the latest completed run
2. Scroll to **Artifacts** → click **"tpb-link-report-..."** to download

---

## 📊 Excel Report — 4 Tabs

| Tab | What's in it |
|-----|-------------|
| 📊 **Summary** | High-level counts — total URLs, issues by type |
| 🔴 **Broken Links** | Every broken URL, source page, status code, error type |
| 🟡 **Redirect Chains** | All redirect hops with full chain shown |
| 📋 **All Results** | Complete log of every URL checked |

### Colour Coding
| Colour | Meaning |
|--------|---------|
| 🔴 Red | 4xx errors (404, 403, etc.) |
| 🟠 Orange | 5xx server errors |
| 🟡 Yellow | Redirect issues / warnings |
| 🟢 Green | Working correctly |

---

## ⚙️ Setup Checklist

- [ ] Create GitHub repository
- [ ] Add `link_checker.py`, `.github/workflows/link-checker.yml`, `requirements.txt`, `README.md`
- [ ] Add 5 GitHub Secrets (Settings → Secrets → Actions):

| Secret | Value |
|--------|-------|
| `EMAIL_FROM` | Gmail address to send FROM |
| `EMAIL_TO` | Email address to receive reports |
| `EMAIL_PASSWORD` | Gmail App Password (16 chars) |
| `SMTP_SERVER` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |

- [ ] Enable GitHub Actions
- [ ] Do a manual test run ✅

---

## 📅 Run Schedule

| Run Type | When | Who triggers |
|----------|------|-------------|
| Automatic | 28th of every month, 8am AEST | Nobody — fully automated |
| Manual | Any time | Click "Run workflow" in Actions tab |
| Reports kept for | 90 days | Auto-deleted by GitHub after |
