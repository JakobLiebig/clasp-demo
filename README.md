# Google Sheets Add-on with clasp

Build Google Sheets add-ons with VS Code instead of the crappy Google UI. Includes API calls (currency rates) and custom HTML sidebar.

## 1. Quickstart (This Project)

```bash
cd clasp-demo
clasp push
```

Open Google Sheets, refresh page, see "Data Analyzer Pro" menu.

**Features:**
- Menu: Fetch Currency Rates (calls API, populates sheet)
- Menu: Launch Dashboard (opens HTML sidebar with live rates)
- Custom functions: `=CONVERTCURRENCY(100, "USD", "EUR")`
- Auto-format, analyze data, remove duplicates, etc.

## 2. clasp Installation & New Project Setup (Windows)

### Install clasp
```powershell
npm install -g @google/clasp
```

### Login to Google
```bash
clasp login
```
Browser opens, authorize clasp with your Google account.

### Create New Project
```bash
mkdir my-sheets-addon
cd my-sheets-addon
clasp create --type sheets --title "My Add-on"
```

Creates `.clasp.json` and opens Apps Script project.

### Create Your Code
```bash
mkdir src
```

**src/Code.gs:**
```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('My Menu')
    .addItem('Click Me', 'showAlert')
    .addToUi();
}

function showAlert() {
  SpreadsheetApp.getUi().alert('Hello from VS Code!');
}
```

**src/Sidebar.html:**
```html
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: Arial;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
    }
    button {
      padding: 10px 20px;
      background: white;
      border: none;
      border-radius: 5px;
      cursor: pointer;
    }
  </style>
</head>
<body>
  <h2>Custom Sidebar</h2>
  <button onclick="google.script.run.showAlert()">Click Me</button>
</body>
</html>
```

**src/appsscript.json:**
```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

### Update .clasp.json
```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "./src"
}
```

### Push to Google
```bash
clasp push
```

### Test
1. Open Google Sheets
2. Create new spreadsheet
3. Refresh page (F5)
4. See your menu

### Edit Workflow
1. Edit files in VS Code
2. `clasp push`
3. Refresh Sheets
4. Done

That's it. Way better than the UI.
