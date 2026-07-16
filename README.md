# PDF Merge Popup

Hotkey-summoned PDF merge utility. Press **Ctrl+Alt+M** anywhere in Windows and a window pops up showing every PDF you currently have open. Drag them (or double-click) into the merge queue, choose "All" or a page range per file, reorder by dragging the grip handle, then **Merge & Save**.

## Setup (dev / testing)

```
pip install pypdf keyboard psutil
python pdf_merge_popup.py
```

The app starts, shows itself once, and then lives in the background. Esc or closing the window just hides it — Ctrl+Alt+M brings it back. Kill it from Task Manager (python.exe / PDFMergePopup.exe) when you want it gone, or Ctrl+C if running from a console.

## Build the exe

```
pip install pyinstaller
pyinstaller --noconsole --onefile --name PDFMergePopup pdf_merge_popup.py
```

Output lands in `dist\PDFMergePopup.exe`. Drop a shortcut in `shell:startup` if you want it running on login.

## If it disappears or the hotkey stops working

The windowed build has no console, so startup and hotkey failures are written to
`pdfmerge_debug.log` beside the executable. Use the **Log** link in the app, or
open that file directly after a failure. The app now also keeps the window open
with a status message if Windows or a security policy prevents the `keyboard`
hook from registering; it does not silently leave you with a hidden window and
no working hotkey.

## How "Open PDFs" detection works

Window titles tell us the *filename* but not the *path*, so resolution happens in tiers:

1. **Window titles** — every visible window with `.pdf` in the title (Acrobat, Bluebeam Revu, Foxit, Edge, etc.) gets its filename extracted.
2. **Process handles** — psutil inspects open file handles on known PDF viewer processes to map filenames → full paths. Works without admin for your own processes.
3. **Recent Items fallback** — `%APPDATA%\Microsoft\Windows\Recent\*.pdf.lnk` shortcuts are parsed for targets, matching by filename.

Anything still unresolved shows grayed out — double-click it to browse to the file manually. There's also **+ Add PDF…** for files that aren't open at all.

Browser-viewed PDFs (Edge/Chrome) often show a title but no local path (they may be streaming from a server or sitting in a temp cache) — those will typically land in the "unresolved" bucket.

## Page range syntax

Per-file, when "Pages:" is selected:

- `3` — single page
- `1-4` — range
- `6-` — page 6 to end
- `-3` — start to page 3
- `2-4, 7, 10-` — any comma-separated combo; order is preserved in the output

The same file can be added to the queue multiple times with different ranges (useful for pulling a cover page to the front, etc.).

## Notes

- The `keyboard` library registers a global hotkey without admin rights on standard Windows setups. If your work machine's policy blocks low-level keyboard hooks, the app still works — just leave the window open or pin the exe to the taskbar instead of using the hotkey.
- Password-protected/encrypted PDFs will fail at merge time with a clear error naming the file.
