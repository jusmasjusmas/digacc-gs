# Slides A11y

A Google Apps Script add-on for Google Slides that checks for and fixes WCAG 2.1 AA / Title II accessibility issues. Runs as an unpublished personal add-on — no Marketplace submission required.

---

## Checks performed

| Check | Auto-fixable |
|---|---|
| Missing presentation title | Yes |
| Missing or empty slide titles (AI-generated if key set) | Yes (with AI key) |
| Duplicate slide titles | Yes |
| Images missing alt text (AI-generated if key set) | Yes (with AI key) |
| Tables missing alt text description | Yes |
| Empty text boxes | Yes (removes them) |
| Text smaller than 18pt | Flag only |
| No speaker notes | Flag only |
| Low-contrast text color | Flag only |

---

## Setup

### 1. Open Apps Script

Go to [script.google.com](https://script.google.com) and click **New project**.

### 2. Create the files

You need two files: `Code.gs` and `Sidebar.html`.

**Code.gs** is created by default. Paste the contents of `Code.gs` into it.

To create `Sidebar.html`:
1. Click the **+** next to Files
2. Choose **HTML**
3. Name it `Sidebar` (no extension)
4. Paste the contents of `Sidebar.html` into it

### 3. Save the project

Give it a name like `Slides A11y`, then press **Cmd+S**.

### 4. Install as a personal add-on

1. Click **Deploy** → **Test deployments**
2. Click **Install** → **Done**

### 5. Open Google Slides

Open any Google Slides file. You should see **Extensions → Accessibility → Check Presentation** in the menu.

> If the menu doesn't appear, refresh the Slides tab.

---

## Usage

1. Open a presentation in Google Slides
2. Click **Extensions → Accessibility → Check Presentation**
3. The sidebar opens — click **Scan**
4. Review the issues list:
   - **Accept** — applies the suggested fix immediately
   - **Mark Reviewed** — acknowledges a manual-review-only issue
   - **Skip** — dismisses the issue
5. Click **Remediate All** to auto-apply every safe fix in one pass
6. Click **Scan** again after fixing to confirm everything is resolved

---

## AI features (optional)

AI-generated alt text for images and AI-generated slide titles require an Anthropic API key.

1. In the sidebar, click the **Settings** tab
2. Paste your `sk-ant-…` key and click **Save Key**
3. The key is stored in your Google account via `PropertiesService` — it never leaves Google's servers except when making requests to the Anthropic API

Without a key, all other checks still run. AI-dependent issues are flagged for manual input instead.

---

## Notes

- This is an unpublished personal add-on — it only appears in your own Google account
- Fixes are applied directly to the open presentation and are immediately undoable with Cmd+Z
- The add-on only reads/writes the active presentation — no other files are accessed
