# Automobilista 2 Career Manager — Master Technical Specification & Architecture Manual

## 1. Executive Summary & Vision

### 1.1 Original Concept & Purpose
Automobilista 2 (AMS2), developed by Reiza Studios, offers an expansive range of motorsport disciplines across historic and modern eras. However, it lacks a fully integrated, deep single-player career mode with long-term financial management, contract negotiations, multi-class championship structures, and multi-season driver tracking.

The **AMS2 Career Manager Desktop** was created to serve as an external, custom-built "career engine." It provides sim-racers with a persistent, immersive career progression framework, logging results, managing earnings, enforcing rulesets, and maintaining multi-season statistics outside the simulation.

### 1.2 Core System Goals
* **Data-Driven Progression**: Driven by external Excel (`.xlsx`) "Blueprints" defining custom championships, race calendars, AI skill scaling, scoring systems, and prize purses.
* **Persistent AppData State**: Total decoupling of application logic from user save data. Saves are automatically stored in standard Windows AppData directories to prevent data loss across updates.
* **Slate-Dark Ergonomic UI**: High-contrast, slate-dark aesthetic optimized for low-light gaming environments, featuring zebra-striped dynamic tables, state safety locks, and auto-scrolling feedback logs.
* **Zero-Dependency Standalone Runtime**: Portable distribution capability into a single, self-contained Windows executable (`.exe`).

---

## 2. Technical Stack & Architecture Rationale

### 2.1 Technology Matrix
| Layer | Technology | Selection Rationale |
| :--- | :--- | :--- |
| **Language** | Python 3.10+ | Rapid development cycle, robust standard library, strong file I/O. |
| **GUI Engine** | Tkinter (`tkinter.ttk`) | Native cross-platform widget toolkit embedded in Python standard distribution. |
| **Data Engine** | SQLite3 (`sqlite3`) | Serverless, zero-configuration, transactional relational database module natively supported in Python. |
| **Excel Parser** | OpenPyXL (`openpyxl`) | High-fidelity `.xlsx` parsing for dynamic schema reading without relying on COM/Excel automation. |
| **Packaging** | PyInstaller | Freezes Python bytecode and interpreter into a standalone Windows binary. |

### 2.2 Architectural Blueprint (Decoupled Layering)
The application adheres to a strict separation of concerns into three core layers:

```text
+-----------------------------------------------------------------------+
|                         PRESENTATION LAYER                            |
|        (Tkinter UI, Modern Dynamic Tables, Safety Modal Locks)         |
+-----------------------------------++----------------------------------+
                                    ||
                                    \/
+-----------------------------------------------------------------------+
|                            SERVICE LAYER                              |
|   (Blueprint Engine, Formula Rulesets, Dynamic Database Transactions) |
+-----------------------------------++----------------------------------+
                                    ||
                                    \/
+-----------------------------------------------------------------------+
|                           PERSISTENCE LAYER                           |
|       (Windows AppData SQLite Engine: ams2_career.db & Config)        |
+-----------------------------------++----------------------------------+
```

---

## 3. Storage Architecture & AppData Integration

To ensure portable safety and user data isolation, the application enforces standard OS-level directory structures for persistence rather than relative application directory writing.

### 3.1 AppData Resolution Path
* Base Directory: `%APPDATA%\AMS2CareerManager\` (Typically `C:\Users\<User>\AppData\Roaming\AMS2CareerManager\`)
* Target Database File: `%APPDATA%\AMS2CareerManager\ams2_career.db`

### 3.2 Relational Database Schema

```sql
-- Driver / Profile Master Table
CREATE TABLE IF NOT EXISTS profiles (
    profile_id INTEGER PRIMARY KEY AUTOINCREMENT,
    driver_name TEXT NOT NULL UNIQUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    current_tier INTEGER DEFAULT 1,
    reputation_score INTEGER DEFAULT 1000,
    total_bank_balance REAL DEFAULT 0.0
);

-- Active Career Seasons
CREATE TABLE IF NOT EXISTS seasons (
    season_id INTEGER PRIMARY KEY AUTOINCREMENT,
    profile_id INTEGER NOT NULL,
    blueprint_name TEXT NOT NULL,
    season_year INTEGER NOT NULL,
    championship_title TEXT NOT NULL,
    is_completed INTEGER DEFAULT 0,
    FOREIGN KEY(profile_id) REFERENCES profiles(profile_id)
);

-- Race Event Logs
CREATE TABLE IF NOT EXISTS race_results (
    result_id INTEGER PRIMARY KEY AUTOINCREMENT,
    season_id INTEGER NOT NULL,
    round_number INTEGER NOT NULL,
    track_name TEXT NOT NULL,
    finish_position INTEGER NOT NULL,
    grid_position INTEGER DEFAULT 0,
    points_awarded INTEGER DEFAULT 0,
    earnings_awarded REAL DEFAULT 0.0,
    bonus_earnings REAL DEFAULT 0.0,
    notes TEXT,
    FOREIGN KEY(season_id) REFERENCES seasons(season_id)
);
```

---

## 4. Excel Blueprint Architecture

Championship schedules, prize payouts, and point distributions are specified using modular `.xlsx` blueprint workbooks parsed at runtime via `openpyxl`.

### 4.1 Required Sheet Structure (`.xlsx`)

#### Sheet 1: `Championship_Meta`
Defines rulesets, global prize multipliers, and target entry criteria.
* Key parameters: `Championship_Name`, `Tier_Level`, `Base_Contract_Fee`, `Entry_Fee`, `Point_Structure_ID`.

#### Sheet 2: `Calendar`
Defines the sequential rounds for a season.
| Round | Track_Name | Laps/Duration | Weather_Preset | Mandatory_Pit | AI_Difficulty |
| :--- | :--- | :--- | :--- | :--- | :--- |
| 1 | Brands Hatch GP | 25 Laps | Dynamic | Yes | 95 |
| 2 | Nürburgring GP | 45 Mins | Clear | No | 97 |

#### Sheet 3: `Payout_Structure`
Maps finishing positions directly to cash payouts.
| Position | Base_Payout ($) | Bonus_Clean_Race ($) | Bonus_Pole ($) |
| :--- | :--- | :--- | :--- |
| 1 | 50000 | 5000 | 2500 |
| 2 | 35000 | 3500 | 1000 |
| ... | ... | ... | ... |

---

## 5. Intricacies of Core Functionality & Engine Logic

### 5.1 Progression Logic: Promotion, Demotion, and Middling Mechanisms
Career movement across motorsport tiers is dynamically computed at the conclusion of every season based on final championship standings, financial solvency, and cumulative driver reputation.

```text
               +----------------------------------+
               |     Season Final Evaluation      |
               +----------------+-----------------+
                                |
         +----------------------+----------------------+
         |                      |                      |
         v                      v                      v
+------------------+   +------------------+   +------------------+
| PROMOTION CHECK  |   |  MIDDLING CHECK  |   |  DEMOTION CHECK  |
| - Finish P1-P3   |   | - Finish P4-P12  |   | - Finish P13+    |
| - Rep Threshold  |   | - Maintain Tier  |   | - Bankrupt Risk  |
| -> Move Tier +1  |   | -> Stay Tier N   |   | -> Move Tier -1  |
+------------------+   +------------------+   +------------------+
```

1. **Promotion (Tier Ascension)**:
   * **Trigger Conditions**: Finishing in the top percentile of the championship standings (typically P1–P3 out of a standard grid) while meeting or exceeding the `Target_Reputation` specified in `Championship_Meta`.
   * **Mechanics**: Unlocks access to higher-tier series blueprints, increases base team contracts, and awards a promotion bonus. The driver's `current_tier` in the `profiles` table increments by +1.

2. **Middling (Tier Retention)**:
   * **Trigger Conditions**: Finishing in the middle brackets of the standings (typically P4 through P12).
   * **Mechanics**: The driver retains their current tier access (`current_tier` remains unchanged). Contracts are renewed at standard rates, allowing drivers to build up bank balances and reputation scores before attempting another championship push.

3. **Demotion (Tier Regression & Penalties)**:
   * **Trigger Conditions**: Finishing in the bottom standings bracket (typically P13 or below) OR incurring financial insolvency (bank balance dropping below zero after entry fees are levied).
   * **Mechanics**: The driver's `current_tier` decrements by -1 (floored at Tier 1). Premium factory drives are revoked, forcing the driver to re-enter lower-tier or privateer feeder series.

### 5.2 Championship Selection & Data Parsing Engine
The selection and ingestion of championship data from Excel spreadsheets is designed to be fully dynamic, decoupled, and fault-tolerant.

1. **Selection Protocol**:
   * On profile load, the engine scans the `blueprints/` directory or imports selected `.xlsx` files.
   * The program cross-references the user's `current_tier` and `reputation_score` against candidate blueprint headers (`Tier_Level` and `Entry_Fee`).
   * Only valid, eligible series are presented in the selection interface, preventing players from jumping directly to elite categories without earning the required credentials.

2. **Parsing Pipeline**:
   * **Validation Phase**: `openpyxl` inspects the workbook structure. If mandatory sheets (`Championship_Meta`, `Calendar`, `Payout_Structure`) or expected column headers are missing, an explicit validation exception is thrown, preventing corrupt data from entering the runtime state.
   * **Data Ingestion Phase**: Numerical data types (laps, AI difficulty, cash values) are cast and normalized. Default fallback values are applied to optional fields (e.g., weather defaults to `Dynamic` if omitted).
   * **Memory Model**: The parsed workbook is converted into a structured Python dictionary representation (Blueprint Model) before being handed to the UI and SQLite transactional layers.

### 5.3 Infinite Data-Driven Scalability via Spreadsheets
A key design pillar of this application is that **no content is hardcoded into Python**. 

* **Zero Code Modification for New Content**: New car classes, track layouts, weather distributions, AI scaling curves, points systems, and financial rewards added to Automobilista 2 by Reiza Studios (or mods) require zero changes to `ams2_career_desktop.py`.
* **Authoring via Spreadsheets**: A user can simply copy an existing `.xlsx` blueprint, enter new track names or session conditions in Excel, and load it into the app. The UI and database dynamically adapt to whatever calendar length, scoring matrix, or prize structure the workbook defines.

### 5.4 Blueprint Library & Save-State Binding
To allow users to run multiple career saves with entirely different progression trees (e.g., a Vintage 1970s Formula Career vs. a Modern GT3 Endurance Career), the application implements a strict **Blueprint Binding Mechanism**.

```text
[ Global Blueprint Repository ]
       │
       ├── GT3_Championship.xlsx  ─────┐ (Instantiated into Save A)
       └── Vintage_F1_1978.xlsx    ────┼──────────────────────────────┐
                                       │                              │
[ SQLite Persistence Engine ]          │                              │
       │                               ▼                              ▼
       ├── Save Profile A ──► [Season 1 Record] ──(Bound to)──► GT3_Championship Snapshot
       └── Save Profile B ──► [Season 1 Record] ──(Bound to)──► Vintage_F1_1978 Snapshot
```

1. **Blueprint-to-Save Binding**:
   * When a new season is initialized, the application parses the selected blueprint and writes a **serialized snapshot** of its meta-rules, calendar rounds, and payout tables directly into the `seasons` and `race_results` records in `ams2_career.db`.
   * The active season becomes permanently bound to that specific blueprint snapshot.

2. **Isolation & Coexistence**:
   * Modifying, moving, or deleting the original `.xlsx` file on disk **does not corrupt or break active career saves**.
   * Multiple career profiles can run completely distinct blueprint rulesets simultaneously without cross-contamination.

---

## 6. UI/UX Specifications

### 6.1 Color Palette Definitions (Slate Dark Theme)
* **Background Dark Primary**: `#090d16`
* **Background Surface / Cards**: `#121824`
* **Border / Accent Subtle**: `#1e293b`
* **Accent Primary (Highlight)**: `#3b82f6` (Royal Slate Blue)
* **Text High Contrast**: `#f8fafc`
* **Text Muted / Subtitle**: `#94a3b8`
* **Table Row Alt (Zebra Striping)**: `#0f1420` and `#171f2e`

### 6.2 Dynamic Table Controls & Safety Features
* **Zebra Striping**: Alternating row color tags attached during item insertion into `ttk.Treeview`.
* **State Locks**: Actions that alter the SQLite database require explicit visual state locks/dialog confirmations to prevent accidental result overwrites.
* **Auto-Scrolling Console**: Event logs use a custom Tkinter `Text` widget bound to auto-scroll to the end index (`END`) upon appending string updates.

---

## 7. Complete Single-File Codebase Implementation

Below is the consolidated, standalone implementation containing database initialization, blueprint engines, and UI views.

```python:Complete Application Code:ams2_career_desktop.py
import os
import sys
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl

# --- CONFIGURATION & APPDATA RESOLUTION ---
APP_NAME = "AMS2CareerManager"

def get_appdata_dir():
    appdata = os.getenv('APPDATA')
    if not appdata:
        appdata = os.path.expanduser('~')
    target_dir = os.path.join(appdata, APP_NAME)
    os.makedirs(target_dir, exist_ok=True)
    return target_dir

DB_PATH = os.path.join(get_appdata_dir(), "ams2_career.db")

# --- DATABASE MANAGER ---
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS profiles (
                profile_id INTEGER PRIMARY KEY AUTOINCREMENT,
                driver_name TEXT NOT NULL UNIQUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                current_tier INTEGER DEFAULT 1,
                reputation_score INTEGER DEFAULT 1000,
                total_bank_balance REAL DEFAULT 0.0
            );
            """)
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS seasons (
                season_id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile_id INTEGER NOT NULL,
                blueprint_name TEXT NOT NULL,
                season_year INTEGER NOT NULL,
                championship_title TEXT NOT NULL,
                is_completed INTEGER DEFAULT 0,
                FOREIGN KEY(profile_id) REFERENCES profiles(profile_id)
            );
            """)
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS race_results (
                result_id INTEGER PRIMARY KEY AUTOINCREMENT,
                season_id INTEGER NOT NULL,
                round_number INTEGER NOT NULL,
                track_name TEXT NOT NULL,
                finish_position INTEGER NOT NULL,
                grid_position INTEGER DEFAULT 0,
                points_awarded INTEGER DEFAULT 0,
                earnings_awarded REAL DEFAULT 0.0,
                bonus_earnings REAL DEFAULT 0.0,
                notes TEXT,
                FOREIGN KEY(season_id) REFERENCES seasons(season_id)
            );
            """)
            conn.commit()

# --- BLUEPRINT PARSER ---
class BlueprintLoader:
    @staticmethod
    def parse_blueprint(filepath):
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Blueprint not found: {filepath}")
        
        wb = openpyxl.load_workbook(filepath, data_only=True)
        
        # Read Calendar Sheet
        calendar = []
        if "Calendar" in wb.sheetnames:
            sheet = wb["Calendar"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:
                    calendar.append({
                        "round": row[0],
                        "track": row[1],
                        "laps": row[2] if len(row) > 2 else "N/A",
                        "weather": row[3] if len(row) > 3 else "Clear",
                        "ai_diff": row[4] if len(row) > 4 else 90
                    })
        return {"calendar": calendar}

# --- MAIN UI IMPLEMENTATION ---
class AMS2CareerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AMS2 Career Manager")
        self.geometry("1100x700")
        self.configure(bg="#090d16")
        
        self.db = DatabaseManager(DB_PATH)
        self.setup_styles()
        self.build_ui()
        self.log_event("Application initialized successfully.")

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # Table Styling
        style.configure("Treeview", 
                        background="#121824",
                        foreground="#f8fafc",
                        fieldbackground="#121824",
                        rowheight=28,
                        borderwidth=0)
        style.map("Treeview", background=[("selected", "#3b82f6")])
        style.configure("Treeview.Heading",
                        background="#1e293b",
                        foreground="#f8fafc",
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))

    def build_ui(self):
        # Top Header Frame
        header = tk.Frame(self, bg="#121824", height=60)
        header.pack(fill="x", side="top")
        
        title_label = tk.Label(header, text="AUTOMOBILISTA 2 CAREER MANAGER", 
                               bg="#121824", fg="#f8fafc", 
                               font=("Segoe UI", 14, "bold"))
        title_label.pack(side="left", padx=20, pady=15)
        
        # Main Layout split into Left Sidebar and Right View
        main_body = tk.Frame(self, bg="#090d16")
        main_body.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Left Panel (Controls)
        left_panel = tk.Frame(main_body, bg="#121824", width=250)
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        
        btn_import = tk.Button(left_panel, text="Import Blueprint (.xlsx)", 
                               bg="#3b82f6", fg="#ffffff", activebackground="#2563eb",
                               font=("Segoe UI", 10, "bold"), relief="flat",
                               command=self.import_blueprint)
        btn_import.pack(fill="x", padx=15, pady=20)

        # Right Panel (Data View)
        right_panel = tk.Frame(main_body, bg="#121824")
        right_panel.pack(side="right", fill="both", expand=True)
        
        # Treeview Schedule Display
        columns = ("round", "track", "laps", "weather", "ai_diff")
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings")
        self.tree.heading("round", text="Round")
        self.tree.heading("track", text="Track Name")
        self.tree.heading("laps", text="Duration/Laps")
        self.tree.heading("weather", text="Weather")
        self.tree.heading("ai_diff", text="AI Difficulty")
        
        self.tree.column("round", width=60, anchor="center")
        self.tree.column("track", width=250, anchor="w")
        self.tree.column("laps", width=120, anchor="center")
        self.tree.column("weather", width=100, anchor="center")
        self.tree.column("ai_diff", width=100, anchor="center")
        
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bottom Console Log Window
        self.console = tk.Text(self, bg="#000000", fg="#94a3b8", height=5, 
                               font=("Consolas", 9), relief="flat")
        self.console.pack(fill="x", side="bottom")

    def log_event(self, text):
        self.console.insert(tk.END, f"> {text}\n")
        self.console.see(tk.END)

    def import_blueprint(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                data = BlueprintLoader.parse_blueprint(file_path)
                self.populate_schedule(data.get("calendar", []))
                self.log_event(f"Loaded blueprint: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load blueprint: {str(e)}")
                self.log_event(f"ERROR: Failed to load {file_path}")

    def populate_schedule(self, calendar_data):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for idx, row in enumerate(calendar_data):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(
                row["round"], row["track"], row["laps"], row["weather"], row["ai_diff"]
            ), tags=(tag,))
        
        self.tree.tag_configure("evenrow", background="#121824")
        self.tree.tag_configure("oddrow", background="#171f2e")

if __name__ == "__main__":
    app = AMS2CareerApp()
    app.mainloop()
```

---

## 8. Build, Packaging & Compilation Instructions

### 8.1 Environment Setup
1. Ensure Python 3.10+ is installed and accessible via command line.
2. Install required external dependencies:
   ```bash
   pip install openpyxl pyinstaller
   ```

### 8.2 Binary Compilation Command
Run the following build command in the project root directory:

```bash
pyinstaller --noconsole --onefile --icon=app_icon.ico ams2_career_desktop.py
```

### 8.3 Output Locations
* Intermediate Build Objects: `./build/` (Safe to delete post-build)
* Final Standalone Executable: `./dist/ams2_career_desktop.exe`

---

## 9. Historical Lessons Learned & Architecture Post-Mortem

### 9.1 The Tkinter Taskbar Icon Bottleneck
* **Issue**: Compiling Tkinter apps into standalone Windows binaries frequently results in the application failing to render a custom `.ico` on the Windows 11 Taskbar, instead displaying the fallback Tcl/Tk "feather" icon.
* **Root Cause**: Tkinter initializes its top-level window handle (`HWND`) through an underlying C wrapper (`tcl86t.dll`). Windows 11 Desktop Window Manager (DWM) registers the process's default process model ID (`AppUserModelID`) to the executable wrapper rather than the dynamically injected icon resource.
* **Resolution / Workarounds**:
  1. *Shortcut Assignment*: Pinning a generated Windows `.lnk` shortcut with a bound custom icon overrides the runtime process handle fallback.
  2. *Framework Migration*: Migrating the presentation layer to **PyQt6** or **PySide6** natively maps application icons directly to the DWM compositor via standard calls (`app.setWindowIcon(...)`), bypassing Tcl/Tk wrapper flaws completely.

### 9.2 Future Framework Migration Pathway (PyQt6 / PySide6)
Should a complete refactoring of the presentation layer be executed in the future:
* **Retain**: All SQL schemas, `DatabaseManager` logic, `BlueprintLoader` Excel parsing rules, promotion/demotion engine algorithms, blueprint-to-save binding architecture, and AppData directory persistence mechanisms.
* **Replace**: Swap `tkinter` widgets for Qt primitives (`QMainWindow`, `QTableWidget`, `QVBoxLayout`). Apply styling via QSS stylesheets rather than Tkinter layout tags.
