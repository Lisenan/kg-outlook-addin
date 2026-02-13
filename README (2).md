# KG Quick Reply & Postings — Outlook Add-in

A task pane add-in for Outlook that gives you one-click reply templates and a load postings manager, built for Kingsgate Logistics.

---

## Features

### ⚡ Quick Replies Tab
- **Built-in templates**: Covered, DNU, Highway — one click to Reply All
- **Custom templates**: Add your own with name, color, and body text
- **Two modes per template**:
  - **Send Now** → Opens Reply All with the template pre-filled (review + send)
  - **Preview Draft** → Same as Send Now but designed for review before sending
- If Outlook isn't connected (dev mode), it falls back to copying text to clipboard

### 📦 Postings Tab
- **Paste & clean** raw load postings (same cleaning logic from your Python script)
- Auto-detects PU/DEL counts and formats the header
- Optional pay rate override
- **Per-posting actions**: Reply All with posting, Copy, Edit, Delete
- Expand/collapse to preview full posting text

### ⚙️ Settings Tab
- Toggle between Reply All and Reply (sender only)
- Custom email signature (appended to all replies)
- Export/Import data as JSON backup
- Reset all data

---

## Quick Start (Local Development)

### Prerequisites
- [Node.js](https://nodejs.org/) v16 or later
- Outlook desktop (Windows/Mac) or Outlook on the web (outlook.office.com)

### Steps

```bash
# 1. Navigate to the add-in folder
cd kg-outlook-addin

# 2. Install dependencies
npm install

# 3. Install HTTPS dev certs + start server
npm run dev

# 4. Server starts at https://localhost:3000
```

### Sideloading into Outlook

#### Option A: Outlook on the Web (Easiest)
1. Go to https://outlook.office.com
2. Open any email
3. Click **"..." (More actions)** → **"Get Add-ins"**
4. Click **"My add-ins"** → **"Add a custom add-in"** → **"Add from file"**
5. Select the `manifest.xml` file from this folder
6. The "KG Quick Reply" button will appear in your ribbon

#### Option B: Outlook Desktop (Windows)
1. In Outlook, go to **File** → **Manage Add-ins** (opens web browser)
2. Click **"My add-ins"** → **"Add a custom add-in"** → **"Add from file"**
3. Select `manifest.xml`
4. Restart Outlook — the button appears in the ribbon

#### Option C: Microsoft 365 Admin Center (Org-wide)
1. Go to https://admin.microsoft.com
2. **Settings** → **Integrated apps** → **Upload custom apps**
3. Upload `manifest.xml`
4. Assign to users/groups
5. Add-in becomes available to everyone in the org

---

## File Structure

```
kg-outlook-addin/
├── manifest.xml        # Outlook add-in manifest (points to taskpane.html)
├── taskpane.html       # Complete add-in UI (HTML + CSS + JS in one file)
├── server.js           # Local HTTPS dev server
├── package.json        # Node.js project config
├── assets/             # Icons (you'll need to add these)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
└── README.md           # This file
```

---

## Deploying to Production

For production, you'll need to host `taskpane.html` on an HTTPS server. Options:

### Option 1: Azure Static Web Apps (Free tier available)
```bash
# Install Azure CLI, then:
az staticwebapp create --name kg-outlook-addin --source .
```
Then update all URLs in `manifest.xml` from `https://localhost:3000` to your Azure URL.

### Option 2: Any HTTPS Web Server
Upload `taskpane.html` and the `assets/` folder to any HTTPS-enabled host (IIS, Nginx, S3 + CloudFront, etc.) and update `manifest.xml` URLs accordingly.

---

## Icons

You'll need to create or provide icon PNG files in the `assets/` folder at these sizes:
- 16x16, 32x32, 64x64, 80x80, 128x128

A simple "KG" logo on a blue (#3b82f6) background works great. You can generate these with any image editor or online tool.

---

## How It Works

| Action | What Happens |
|--------|-------------|
| Click **⚡ Send Now** | Opens Outlook Reply All form pre-filled with template text |
| Click **👁 Preview Draft** | Same — opens reply form for you to review before sending |
| Click **📋** on a posting | Copies posting text to clipboard |
| Click **⚡** on a posting | Opens Reply All with the posting content |

> **Note**: The Office.js API opens reply forms but cannot auto-send emails without additional Exchange admin permissions. Both "Send Now" and "Preview" open the form — you hit Send in Outlook.

---

## Data Storage

All data (custom templates, postings, settings) is stored in **browser localStorage**. This means:
- Data persists across sessions
- Data is per-browser (won't sync between devices)
- Use **Export/Import** in Settings to back up or transfer data
- If you clear browser data, your templates and postings will be lost (export first!)

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Add-in doesn't appear in ribbon | Make sure the dev server is running at https://localhost:3000 |
| Certificate error in browser | Run `npm run dev` to install trusted dev certs |
| "Reply form unavailable" toast | Make sure you have an email selected/open in Outlook |
| Templates not saving | Check that localStorage isn't disabled in your browser |
