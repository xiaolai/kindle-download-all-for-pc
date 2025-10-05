# Kindle Download Automator for Windows

Automation script that iterates through your Kindle for PC library and triggers the **Download** command on each title. It uses `pywinauto` to drive the Kindle UI, starting from either the currently selected book or a 1-based index supplied at runtime.

## Prerequisites

- Windows with Kindle for PC installed and signed in.
- Python 3.12 (or compatible) with the following packages installed in the same environment:
  - `pywinauto`
  - `pywin32`
  - `six`

Install dependencies with:

```powershell
python -m pip install pywinauto
```

> **Tip:** Launch Kindle for PC before running the script and leave the library view in focus.

## Files

- `kindle_download_all.py` – Automation script entry point.
- `README.md` – This guide.

## Usage

1. Ensure Kindle for PC is running and the library window is visible.
2. Optionally, click the book you want the automation to start with.
3. From a PowerShell prompt in this folder, run:

   ```powershell
   python kindle_download_all.py [--start-from N] [--max-iterations M]
   ```

   - `--start-from N` (optional): Start at the Nth book in the Kindle list (1-based). If omitted, the script starts from the currently focused/selected item.
   - `--max-iterations M` (optional): Limit how many books are processed in one run (default: 10000).

The script logs each title as it moves through the list, invoking the `context-menu-option-download-book` menu item when available, or sending Down/Down/Enter as a fallback.

## Behaviour Details

- Uses index-based iteration, so it progresses reliably even when Kindle lazily loads titles.
- Wait timings are controlled by constants near the top of the script (`MENU_APPEAR_TIMEOUT`, `DELAY_BETWEEN_ACTIONS`, etc.). Increase them if your system responds slowly.
- Stops automatically when Kindle no longer returns new items or when the iteration cap is hit.

## Safety Notes

- The script performs UI automation; avoid interacting with the mouse/keyboard while it is running.
- Always test on a small range (e.g., `--max-iterations 5`) before processing your entire library.

## Troubleshooting

- **“Process "Kindle.exe" not found!”** – Launch Kindle for PC first.
- **Context menu doesn’t appear** – Increase `MENU_APPEAR_TIMEOUT` and verify the Kindle window has focus.
- **Downloads start partway through the list** – Use `--start-from` with the desired index or manually select the correct starting book before running.

Feel free to adjust the script for your workflows; contributions and further automation ideas are welcome!
\nTested with Kindle for PC version 2.8.0 (70980), which is compatible with Epubor Ultimate 3.0.16.508 for post-download processing.
