# 🚚 TruckTalk Connect — Google Sheets Add-on

Analyze logistics load sheets directly inside Google Sheets.  
The add-on flags issues (missing columns, duplicate IDs, bad dates, etc.) and produces Loads JSON ready for TruckTalk.

---

## Setup

### Requirements
- Google account with access to Apps Script
- OpenAI API key
- (Optional) Node.js + CLASP if you want to sync the code locally

### Installation

### Option A — Quick Setup (copy-paste into Google Sheets)
1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Paste the contents of `code.gs`, `ui.html`, and `appsscript.json` into the Apps Script editor
4. Set your OpenAI API key:
   - In Apps Script editor → **Project Settings > Script Properties**
   - Add property `OPENAI_API_KEY` with your key
5. Save and reload the Sheet
6. You should see a new menu: **TruckTalk Connect**


### Option B — Developer Setup (CLASP + VSCode)

1. Clone the repo
   
   git clone <repository_url>
   cd trucktalk-connect

2. Install CLASP (Apps Script CLI)

   npm install -g @google/clasp
   clasp login

3. Create or link an Apps Script project

   clasp create --title "TruckTalk Connect" --type sheets
   
   or if you already have a project:
   clasp clone <scriptId>

4. Push local code to Apps Script

   clasp push

5. Open the linked Google Sheet → reload → menu TruckTalk Connect appears.

---

## Usage

1. Open your sheet with loads data (headers like `Load ID`, `PU Date`, etc.)
2. Go to menu **TruckTalk Connect > Open Sidebar**
3. In the sidebar (chatbox):
   - Type “analyze”, “Scan Tab”, "review this tab" etc or click the Re-analyze button 
   - Issues will be flagged in **Results tab**
   - Apply suggested fixes if available
4. After re-analysis, **JSON output** appears in the Results tab for export

---

## Testing

- Sample data: [Loads Sample Sheet](https://docs.google.com/spreadsheets/d/1uAnHCDwm3847CiYFXIvYj4KTeSUtZdrxabMsNS0cUjI/edit?usp=sharing)
- Include:
  - ✅ 2 happy rows (no issues)
  - ❌ 3 broken rows (e.g., missing columns, bad dates, duplicate ID)
- Manual test flow:
  - Analyze → Review issues → Fix → Re-analyze → JSON output
- Unit tests (in `code.gs`):
  - Pure utilities 
  - Run them inside Apps Script

---

## Deliverables

- Repo: contains `code.gs`, `ui.html`, `appsscript.json`, and this README
- Screencast (≤ 2 min): show
  1. Analyze a sheet
  2. See issues
  3. Apply a fix
  4. Show JSON output

Screencast_link = https://drive.google.com/file/d/1hHi23qyCO1BXRyoMOHwkv8dl2xmZmUDu/view?usp=sharing
---

## Limitations

- Date normalization is partial (some formats may not convert perfectly)
- JSON generation depends on if AI able to fix the issues or not
- AI responses may vary from every analysis (maybe right or wrong)
- Auto-fixes are basic — user confirmation still required



