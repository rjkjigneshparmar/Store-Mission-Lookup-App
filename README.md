# Store Mission Lookup

<img width="1280" height="640" alt="social_preview_1" src="https://github.com/user-attachments/assets/0cf5fd60-948c-4825-b1e8-02b2c42e4923" />

A Windows desktop application for field operations teams to manage store-mission (survey) assignments. When a store permanently closes, instantly find every mission it belongs to and remove it — individually, in bulk, or all at once.

Built with **Python**, **customtkinter**, and **SQLite**.

---

## The Problem

When a store closes, teams had no quick way to see which missions (surveys) it was assigned to. Checking each mission manually across hundreds of campaigns was time-consuming and error-prone.

**Store Mission Lookup solves this** — import your store and mission data, search any store by name, and see all its missions instantly.

---

## Features

- **CSV & Excel import** — supports `.csv`, `.txt`, `.xlsx`, and `.xls`
- **3-stage validation** — file type check, column detection, and missing data pre-scan before anything is written to the database
- **Fuzzy search** — finds stores even with typos (e.g. `hannford` → `Hannaford #8327`)
- **Bulk delete + Undo** — delete any number of stores at once with a 30-second undo window
- **Shared team database** — point all team members to the same file on OneDrive, SharePoint, or Google Drive
- **Who's online** — see which teammates have the app open in real time
- **Change notifications** — get alerted when a teammate imports or deletes data
- **Import lock** — prevents two people writing to the database at the same time
- **Activity history** — full audit trail of every import, delete, and undo with sortable, resizable columns
- **Import DB file** — merge or replace data from another machine's database
- **VACUUM on delete** — database file shrinks on disk after deletions
- **Auto-refresh** — background refresh every 10–45 minutes for shared database use
- **Keyboard shortcuts** — full shortcut set with a built-in reference window (F1)

---

## Screenshots

<img width="1798" height="910" alt="Screenshot 2026-05-28 173927" src="https://github.com/user-attachments/assets/7e494140-825b-437b-b800-1e31fb3df1c7" />

---

## Requirements

- Windows 10 or Windows 11
- Python 3.8 or higher
- Internet connection (first-time setup only, to install dependencies)

---

## Installation

### Option A — Build a standalone .exe (recommended)

No Python required on the target machine after building.

1. Clone or download this repository
2. Place `app_v4.py` and `build_v4.bat` in the same folder
3. Double-click `build_v4.bat`
4. Wait 1–3 minutes while dependencies are installed and the .exe is built
5. Your app is ready at `dist\StoreMissionLookup_v4.exe`
6. Move the `.exe` anywhere — it works from any location

> **First launch:** Windows may show a SmartScreen warning. Click **More info** → **Run anyway**. This only happens once.

### Option B — Run directly with Python

```bash
# Install dependencies (first time only)
pip install customtkinter openpyxl

# Run the app
python app_v4.py
```

---

## CSV / Excel File Format

Your import file must have exactly **3 columns**. Every row must have all 3 cells filled in — blank cells will block the import entirely.

| Column | Description | Example |
|---|---|---|
| `Store Details` | Full store name and address | `Hannaford #8327 - 777 Rogers St, Lowell, MA 01852` |
| `Mission Name` | Name of the survey or mission | `Find the honey` |
| `Mission Link` | Direct URL to the mission | `https://admin.jignesh.app/#/campaigns/2elmBuT5FG/show` |

- One row per store–mission pair
- If a store is in 3 missions, it appears on 3 separate rows
- Duplicate rows are skipped automatically — safe to re-import the same file

---

## Shared Database Setup (Team Use)

To share data across a team, point everyone's app to the same database file on a synced folder.

**Step 1 — One person (admin) sets up the shared folder:**

| Platform | Example path |
|---|---|
| OneDrive | `C:\Users\YourName\OneDrive\StoreMissionLookup\` |
| SharePoint (synced) | `C:\Users\YourName\CompanyName\StoreMissionLookup\` |
| Google Drive | `G:\My Drive\StoreMissionLookup\` |

**Step 2 — Share the folder** with all team members via OneDrive/SharePoint sharing or Google Drive sharing so it syncs to their machines.

**Step 3 — Each team member opens Settings (Ctrl+S) and sets:**
```
C:\Users\[TheirName]\OneDrive\StoreMissionLookup\data.db
```

**Step 4 — Click Test Connection → Save → Refresh**

The status bar will show `🔗 Shared` when connected.

> **Team rule:** Only one person imports or deletes at a time. The app shows a 🔒 lock when someone is importing. Everyone else clicks ↻ Refresh to see the latest data.

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl + F` | Focus the search box |
| `Ctrl + R` | Refresh — clears search, filter, and selection |
| `Ctrl + I` | Open Import dialog |
| `Ctrl + E` | Export filtered data to CSV |
| `Ctrl + Z` | Undo last delete (30-second window) |
| `Delete` | Delete selected stores |
| `Ctrl + H` | Open Activity History |
| `Ctrl + S` | Open Settings |
| `Ctrl + \`` | Cycle between open popup windows |
| `F1` | Open Keyboard Shortcuts reference |
| `Escape` | Clear search / close any popup |

---

## Data Storage

Data is stored in a local SQLite database:

```
%APPDATA%\StoreMissionLookup\data.db
```

Which typically resolves to:
```
C:\Users\YourName\AppData\Roaming\StoreMissionLookup\data.db
```

Your data persists between sessions and is unaffected by moving or deleting the `.exe`.

**Backup:** Copy the `StoreMissionLookup` folder from AppData to an external drive or cloud location.

**Reset settings** (if the app crashes on startup):
```bash
del "%APPDATA%\StoreMissionLookup\settings.json"
```

---

## Version History

| Version | Highlights |
|---|---|
| v1.0 | CSV import, local SQLite database, search, export |
| v2.0 | Shared database, multi-select delete, import history, Settings |
| v2.1 | Bulk delete chunking (fixes crash >999 stores), path validation fix |
| v3.0 | Undo delete, who's online, change notifications, import lock, keyboard shortcuts, sortable history, auto-refresh, window cycling |
| v4.0 | Excel support, 3-stage validation, fuzzy search, resizable history columns, VACUUM on delete, Import DB file, dark mode permanent, window auto-centres at 720px |

---

## Tech Stack

| Component | Technology |
|---|---|
| UI framework | [customtkinter](https://github.com/TomSchimansky/CustomTkinter) |
| Database | SQLite (via Python `sqlite3`) |
| Excel support | [openpyxl](https://openpyxl.readthedocs.io/) (`.xlsx`), [xlrd](https://xlrd.readthedocs.io/) (`.xls`) |
| Packaging | [PyInstaller](https://pyinstaller.org/) |
| Platform | Windows 10 / 11 |

---

## License

This project is for internal use. See `LICENSE` for details.
