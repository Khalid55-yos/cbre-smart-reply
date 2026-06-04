# CBRE Smart Reply — Outlook Add-in v3.0

AI-powered listing reply for the LA North Capital Markets team. Adds a "Smart Reply" button next to Reply/Reply All on incoming emails. Click → auto-detects the listing → opens a reply pre-populated with the body and the OM attached.

---

## Deployment Path

There are 4 phases. Total time: ~1.5 hours focused work.

| Phase | What | Estimated time |
|---|---|---|
| 1 | Build the HTTP-triggered Power Automate flow | 30 min |
| 2 | Host these files on GitHub Pages | 20 min |
| 3 | Configure manifest.xml with your URLs | 5 min |
| 4 | Sideload the add-in in Outlook + test | 15 min |

---

## Phase 1 — Build the HTTP-triggered Power Automate Flow

This flow is what the Add-in calls. It receives an email's subject + body, finds the matching listing in SharePoint LISTINGS folder, returns the reply body + OM file as base64.

### Step 1.1 — Create a new flow

1. Go to **https://make.powerautomate.com**
2. Make sure you're in your Developer environment (top-right)
3. Click **+ Create** → **Instant cloud flow**
4. Name: `Smart Reply HTTP Endpoint`
5. Trigger: **When an HTTP request is received**
6. Click **Create**

### Step 1.2 — Configure the trigger schema

Paste this as the **Request Body JSON Schema**:

```json
{
  "type": "object",
  "properties": {
    "subject": { "type": "string" },
    "body": { "type": "string" },
    "senderEmail": { "type": "string" },
    "senderName": { "type": "string" },
    "listingNameOverride": { "type": "string" }
  }
}
```

### Step 1.3 — Add SharePoint action: list all OMs

Add action → **SharePoint → Get files (properties only)**:
- **Site Address:** `https://cbre.sharepoint.com/teams/engineeredforexcellence`
- **Library Name:** Documents
- **Folder Identifier:** browse to `LISTINGS` folder

### Step 1.4 — Add a Filter array to find matching listing

Add action → **Data Operations → Filter array**:
- **From:** the `value` array from "Get files" step
- **Condition:** custom OData filter. Use this expression in advanced mode:

```
@or(
  contains(toLower(item()?['Name']), toLower(triggerBody()?['listingNameOverride'])),
  and(
    not(empty(triggerBody()?['subject'])),
    contains(toLower(item()?['Name']), toLower(first(split(triggerBody()?['subject'], ' '))))
  )
)
```

This filters files where the filename matches either the explicit override OR the first significant word of the email subject. *You'll likely want to refine this matching logic over time.*

### Step 1.5 — Get file content for first match

Add action → **SharePoint → Get file content**:
- **Site Address:** same as above
- **File Identifier:** `first(body('Filter_array'))?['{Identifier}']` — use the expression editor

### Step 1.6 — Generate the reply body

Add a **Compose** action with a templated reply:

```
Hi,

Thank you for your interest in @{first(body('Filter_array'))?['DisplayName']}. Please find the offering memorandum attached for your review.

Happy to set up a call at your convenience to walk through the opportunity in more detail and answer any underwriting questions.

Best,
```

*Note: Outlook will append your signature; do not add one here.*

### Step 1.7 — Get the list of all active listings (for the dropdown override)

Add a **Select** action:
- **From:** `body('Get_files_(properties_only)')?['value']`
- **Map:** create an object with two fields:
  - `name` → `item()?['DisplayName']`
  - `filename` → `item()?['Name']`

### Step 1.8 — Build the response

Add **Response** action (HTTP Response):
- **Status Code:** 200
- **Headers:**
  - `Content-Type`: `application/json`
- **Body:**

```json
{
  "detected": "@{not(empty(body('Filter_array')))}",
  "confidence": "@{if(not(empty(body('Filter_array'))), 'high', 'none')}",
  "listingName": "@{first(body('Filter_array'))?['DisplayName']}",
  "omFilename": "@{first(body('Filter_array'))?['Name']}",
  "replyBody": "@{outputs('Compose')}",
  "omBase64": "@{body('Get_file_content')?['$content']}",
  "allListings": "@{body('Select')}"
}
```

### Step 1.9 — Save and copy the HTTP URL

1. **Save** the flow
2. Open the trigger card again
3. Copy the **HTTP POST URL** that's now displayed
4. Save this URL — you'll paste it into `taskpane.js`

⚠️ This URL contains a SAS token that authenticates the request. Treat it like a secret. Anyone with the URL can call your flow.

---

## Phase 2 — Host on GitHub Pages

### Step 2.1 — Create GitHub account (if needed)

If you don't have GitHub:
1. Go to **https://github.com/signup**
2. Sign up (free)
3. Verify your email

### Step 2.2 — Create a new public repository

1. Click the **+** in the top-right → **New repository**
2. Repository name: `cbre-smart-reply`
3. Description: `CBRE Smart Reply Outlook Add-in`
4. Set as **Public** (required for GitHub Pages on the free tier)
5. **Do NOT** initialize with README (we have our own)
6. Click **Create repository**

### Step 2.3 — Upload these files

1. On the new empty repo page, click **uploading an existing file** (the link)
2. Drag and drop **everything** from this Smart Reply package:
   - `manifest.xml`
   - `commands.html`
   - `taskpane.html`
   - `taskpane.js`
   - `styles.css`
   - `README.md`
   - The entire `icons/` folder (all 5 PNG files)
3. Scroll down → **Commit changes**

### Step 2.4 — Enable GitHub Pages

1. In your repo, click **Settings** (top tab)
2. Left sidebar → **Pages**
3. **Source:** Deploy from a branch
4. **Branch:** `main` (or `master`), folder `/ (root)`
5. Click **Save**
6. Wait ~30 seconds. Refresh the page. You'll see: "Your site is live at https://**YOUR-USERNAME**.github.io/cbre-smart-reply/"
7. Copy that base URL

### Step 2.5 — Verify

Open in a browser: `https://YOUR-USERNAME.github.io/cbre-smart-reply/icons/icon-80.png`

You should see the green star icon. If you see a 404, wait 2 more minutes and try again — GitHub Pages takes a moment to deploy.

---

## Phase 3 — Configure manifest.xml + taskpane.js

You need to update two files with your actual values.

### Step 3.1 — Update manifest.xml

Replace **every occurrence** of `{GITHUB_USERNAME}` in `manifest.xml` with your actual GitHub username.

Example: if your username is `kyosufzai`, then `https://{GITHUB_USERNAME}.github.io/cbre-smart-reply/icons/icon-64.png` becomes `https://kyosufzai.github.io/cbre-smart-reply/icons/icon-64.png`.

There are roughly 9 places to replace. Use Find & Replace (Ctrl+H in most text editors) to make this easy.

### Step 3.2 — Update taskpane.js

At the top of `taskpane.js`, replace:

```
const FLOW_URL = 'PASTE_YOUR_POWER_AUTOMATE_HTTP_TRIGGER_URL_HERE';
```

…with your actual Power Automate HTTP trigger URL from Step 1.9.

### Step 3.3 — Re-upload to GitHub

After editing both files:
1. Go to your GitHub repo
2. Upload the modified `manifest.xml` and `taskpane.js` (drag and drop, will overwrite)
3. Commit changes
4. Wait ~30 seconds for GitHub Pages to redeploy

---

## Phase 4 — Sideload the Add-in in Outlook

### Step 4.1 — Open Outlook on the web

Go to **https://outlook.office.com** and sign in.

### Step 4.2 — Open the Add-ins management

1. Click the **gear icon** (Settings) in the top-right
2. **View all Outlook settings** (bottom of the panel)
3. Left sidebar → **General** → **Manage add-ins**

OR navigate directly to: **https://outlook.office.com/owa/?path=/options/manageapps**

### Step 4.3 — Add a custom add-in

1. Click **+ Add a custom add-in**
2. Choose **Add from URL…** OR **Add from file…**

**Option A: Add from URL** (easier)
- Paste: `https://YOUR-USERNAME.github.io/cbre-smart-reply/manifest.xml`
- Click **OK**

**Option B: Add from file** (if URL doesn't work)
- Download `manifest.xml` from your GitHub repo to your computer
- Click **Add from file…** → choose the downloaded file → OK

### Step 4.4 — Confirm and trust

A confirmation dialog appears. Click **Install** to confirm.

### Step 4.5 — Test it

1. Open any email in your Inbox
2. Look at the message ribbon (above the email body)
3. You should see the **Smart Reply** button with the green star icon, next to Reply/Reply All/Forward

OR for the New Outlook on the web:
- Look at the toolbar
- May be in the **... (More)** menu initially
- Pin to ribbon: right-click → Pin

4. Click **Smart Reply**
5. The taskpane opens on the right
6. It reads the email, calls your flow, and displays the detected listing + preview
7. Click **Open Reply with OM Attached**
8. A reply form opens with the body and OM attached — review, edit, send.

---

## Troubleshooting

### Add-in won't install

- **Manifest validation error:** the {GITHUB_USERNAME} placeholders are still there. Replace them all.
- **CORS error:** GitHub Pages should handle this. If it persists, try sideloading via "Add from file" instead.
- **CBRE tenant policy blocks custom add-ins:** would require IT involvement. There's no workaround at the user level.

### Add-in installs but button doesn't appear

- Hard refresh Outlook (Ctrl+Shift+R)
- Check the **... (More)** menu — Smart Reply may be there
- Right-click an existing ribbon button → look for "Customize ribbon" → enable Smart Reply if disabled

### Button appears but taskpane is empty / shows error

- **"Flow URL not configured":** taskpane.js still has the placeholder URL. Update it and re-upload to GitHub.
- **CORS error from flow:** Power Automate HTTP trigger should allow CORS from your GitHub Pages domain by default. If it doesn't:
  - Open the flow → trigger → there may be a CORS settings advanced option
  - Or use a JSONP wrapper (complex; talk to me)
- **No listing detected:** the filter array logic needs refinement. The simple version uses first-word matching which won't catch every email.

### Attachment doesn't appear in the reply

- The base64 file content might be malformed or too large
- Office.js attachment URL field has a max size; very large OMs (>30MB) may fail
- Try with a smaller test OM first to verify the pipeline works

### Flow errors

- Check the flow run history in Power Automate
- Most common: SharePoint connection isn't authenticated → re-authenticate the connection
- Filter array returning empty → the listing name in the email doesn't match any OM filename

---

## What's in This Package

```
cbre-smart-reply/
├── manifest.xml          The add-in declaration (where to place button, what icons, where source lives)
├── commands.html         Required boilerplate for Office Add-ins
├── taskpane.html         The UI panel that opens when the button is clicked
├── taskpane.js           Logic: reads email, calls flow, opens reply with attachment
├── styles.css            Visual styling for the panel
├── README.md             This file
└── icons/
    ├── icon-16.png       Small ribbon icon
    ├── icon-32.png       Medium ribbon icon
    ├── icon-64.png       Large display icon
    ├── icon-80.png       High-DPI ribbon icon
    └── icon-128.png      Add-in catalog icon
```

---

## What This Doesn't Do (Yet)

- **Real listing auto-detection beyond first-word matching.** The Power Automate filter array is a simple keyword match. For better accuracy, we'd add an LLM call (Azure OpenAI or call your Copilot Studio agent via Direct Line API).
- **Reply body customization per email context.** The body is a generic template. Future: pass email context to an LLM, generate situation-specific replies.
- **Multi-listing replies.** If an inquiry covers 2-3 listings, this attaches one OM. Future: detect multiple, allow multi-attach.
- **Team-wide deployment.** This is personal sideload only. For team-wide, CBRE IT uploads to M365 Integrated Apps.
- **Reply tracking / analytics.** Future: log every Smart Reply use to a SharePoint list for analytics.

---

## Support

If something breaks, the best diagnostics:
1. **Power Automate flow run history** — shows what the flow received and returned
2. **Browser console** in the Outlook taskpane (F12 in some Outlook versions) — shows JavaScript errors
3. **Manifest validator:** https://github.com/OfficeDev/office-addin-validator

---

Built for CBRE LA North Capital Markets | Private Capital Partners
