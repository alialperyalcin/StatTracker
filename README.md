# StatTracker OCR

Desktop app for extracting profile stats from your game window and appending each profile as a new row in Excel.

## What It Does

- Captures the active game window.
- Detects the expected profile stats layout.
- Extracts numeric stats with OCR.
- Appends one row per profile to a `.xlsx` file.
- Supports fast multi-profile capture with hotkeys.

## Requirements

- Windows
- Python 3.10+
- Tesseract OCR installed
- Game profiles shown in the same layout format

Install Tesseract (Windows build):

- https://github.com/UB-Mannheim/tesseract/wiki

Default path used by app:

- `C:\Program Files\Tesseract-OCR\tesseract.exe`

## Installation

Run in project folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Start the App

```powershell
python app.py
```

## Recommended Workflow (Multi-Profile Session)

This is the fastest mode when you open many profiles one after another.

1. Click `Choose Excel File` and select/create your output `.xlsx`.
2. Optional: set `Game Window Title Contains` (part of your game window title) to avoid capturing the wrong window.
3. Optional: set `Tesseract Path` if your install is not in the default location.
4. Click `Start Session`.
5. Open a profile in-game and press `F8`.
6. Repeat step 5 for each profile.
7. Press `F9` or click `Stop Session` when done.

Each `F8` press:

- captures the active window
- extracts stats
- appends one row to Excel

## Single Capture Modes

Use these if you do not want hotkeys.

### Auto Capture + Save

1. Select Excel file.
2. Open profile in game.
3. Click `Auto Capture + Save`.
4. App waits 3 seconds, captures screen, extracts, and saves.

### Manual Screenshot Mode

1. Click `Choose Screenshot`.
2. Click `Extract Stats`.
3. Review/edit values.
4. Click `Save Row to Excel`.

## Output

- Excel rows include:
  - timestamp
  - captured image filename
  - extracted stat fields
- Captured images are stored in `captures/` in the project folder.

## Troubleshooting

### `Permission denied` when saving Excel

Cause: the `.xlsx` file is open/locked (Excel or cloud sync).

Fix:

1. Close the Excel file.
2. Wait a few seconds if OneDrive/Dropbox is syncing.
3. Press `F8` again.

### OCR not working

1. Verify Tesseract is installed.
2. In the app, set `Tesseract Path` explicitly (example: `C:\Program Files\Tesseract-OCR\tesseract.exe`).
3. Retry extraction.

### Wrong window captured

1. Fill `Game Window Title Contains` with a unique part of your game title.
2. Keep game window focused before pressing `F8`.
