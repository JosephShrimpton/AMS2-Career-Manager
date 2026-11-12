# Automobilista 2 Career Manager — Master Technical Specification & Architecture Manual (Version 2.0 & Version 2.3)

## 1. Executive Summary & Vision

### 1.1 Evolution from v1 to v2 & v2.3

Automobilista 2 (AMS2), developed by Reiza Studios, offers an expansive range of motorsport disciplines across historic and modern eras. While version 1 established a foundational external calendar tracking and basic spreadsheet ingestion framework, **Version 2** transformed the application into a robust, telemetry-integrated career engine, and **Version 2.3** elevates it further into a self-contained blueprint authoring suite and non-linear performance evaluation platform.

v2 introduced direct synchronization with **Second Monitor** XML data exports, multi-class championship scoring calculations, persistent manual race result overrides, dynamic contract market evaluations driven by a Trajectory Index (TI), and a complete multi-season historical archive. Building directly atop v2's foundation, **v2.3** introduces:

* An **In-App Blueprint Creator & Editor** (`BlueprintEditorWindow`) with 4-sheet `.xlsx` export, inline double-click championship editing, custom track builders, and automated population of 40 authentic real-team car pairings per newly created championship.

* A dual-mode **Advanced Non-Linear Trajectory Index (TI) & Recovery Engine** featuring non-linear normalization exponents ($1.25$ for championship and race positions, $1.1$ for qualifying), weighted weekend scoring ($40\%$ Championship / $35\%$ Race / $25\%$ Qualifying), and a field-percentage recovery bonus/penalty curve, paired with a live UI toggle to switch between *Advanced Ti* and *Legacy Ti*.

* A **Strict Category Matching Engine** across blueprint auto-population, contract market generation, and initial career setup to eliminate shorthand category collisions (e.g., preventing `"P1"` from matching `"LMP1"`).

* **Save Blueprint Re-assignment & Content Summary Inspection** (`change_save_blueprint()`), enabling dynamic re-binding of `.xlsx` blueprints to active profiles with an interactive summary modal.

* **Master Categories Sheet Ingestion** (`categories` sheet Column A parsing) and expanded dashboard metadata strips (explicitly displaying `Parc Fermé`, active `Classes`, and bound `Blueprint`).

### 1.2 Core System Goals (v2 & v2.3)

* **Telemetry & XML Bridge**: Direct parsing and ingestion of Second Monitor `Championships.xml` files, enabling live standings synchronization, session-by-session classification inspection, and automated qualifying position detection.

* **Multiclass & Scoring Sovereignty**: Dynamic calculation of independent class standings and points allocation based on custom scoring matrices defined within championships, coupled with UI scope filters for both Overall Standings and Per-Round Classifications.

* **Persistent Override & Edit Pipeline**: A robust JSON-based override manager allowing granular editing of race results, finishing positions, class positions, points gains, and fastest laps with strict serialization safeguards.

* **In-App Blueprint Authoring & 4-Sheet Workbook Engine**: Visual creation, modification, staging, and `.xlsx` export of championships, track calendars, car/team rosters, and master vehicle categories directly within the desktop application without requiring external spreadsheet editors.

* **Advanced Market & Progression Continuum**: Algorithmic career evaluation weighting championship standing, non-linear race finish performance, qualifying pace, field-percentage recovery deltas, and field size to generate realistic promotion, retention, lateral move, fallback, and demotion contract offers across strict team/car categories.

* **Zero-Dependency Standalone Runtime**: Portable distribution capability into a single, self-contained Windows executable (`.exe`) utilizing PyInstaller.

## 2. Technical Stack & Architecture Rationale

### 2.1 Technology Matrix

| Layer | Technology | Selection Rationale | 
| ----- | ----- | ----- | 
| **Language** | Python 3.10+ | Rapid development cycle, robust standard library, strong file I/O. | 
| **GUI Engine** | Tkinter (`tkinter.ttk`) | Native cross-platform widget toolkit embedded in Python standard distribution. | 
| **Data Engine** | SQLite3 (`sqlite3`) | Serverless, zero-configuration, transactional relational database module natively supported in Python. | 
| **Excel Parser & Writer** | OpenPyXL (`openpyxl`) | High-fidelity `.xlsx` parsing and multi-sheet workbook generation (`Workbook()`, `load_workbook()`) without relying on COM/Excel automation. | 
| **XML Parser** | ElementTree (`xml.etree.ElementTree`) | Built-in hierarchical XML parsing engine for efficient DOM traversal of Second Monitor exports. | 
| **Packaging** | PyInstaller | Freezes Python bytecode and interpreter into a standalone Windows binary. | 

### 2.2 Architectural Blueprint (Decoupled Layering v2.3)

```
+-----------------------------------------------------------------------+
|                         PRESENTATION LAYER                            |
|   (Tkinter UI, 4 Core Screens, Modern Trees, Modal Round & Blueprint  |
|             Editors, Live Ti Mode Toggle, Content Viewers)            |
+-----------------------------------++----------------------------------+
                                    ||
                                    \/
+-----------------------------------------------------------------------+
|                            SERVICE LAYER                              |
|  (SecondMonitorAdvancedParser, OverrideManager, BlueprintEngine,      |
|   BlueprintEditorWindow Staging, Advanced Non-Linear TI Algorithm)    |
+-----------------------------------++----------------------------------+
                                    ||
                                    \/
+-----------------------------------------------------------------------+
|                           PERSISTENCE LAYER                           |
|  (AppData SQLite Engine: ams2_careers.db, blueprints/*.xlsx, &        |
|                       round_overrides/*.json)                         |
+-----------------------------------------------------------------------+

```

## 3. Storage Architecture & AppData Integration

To ensure portable safety and user data isolation, v2 and v2.3 enforce standard OS-level directory structures for persistence.

### 3.1 AppData Resolution Path

* Base Directory: `%APPDATA%\AMS2CareerManager\`

* Target Database File: `%APPDATA%\AMS2CareerManager\ams2_careers.db`

* Blueprints Library Directory: `%APPDATA%\AMS2CareerManager\blueprints\`

* Round Overrides Directory: `%APPDATA%\AMS2CareerManager\round_overrides\`

### 3.2 Relational Database Schema (`ams2_careers.db`)

```
-- Career Profiles Master Table
CREATE TABLE IF NOT EXISTS careers (
    id TEXT PRIMARY KEY,
    name TEXT,
    driver_name TEXT,
    season INTEGER,
    current_contract TEXT, -- JSON blob
    races TEXT,            -- JSON blob
    calendar TEXT,         -- JSON blob
    history TEXT,          -- JSON blob
    blueprint_name TEXT DEFAULT '',
    second_monitor_id TEXT DEFAULT ''
);

```

*Note: `init_db()` performs non-destructive schema introspection via `PRAGMA table_info(careers)` and executes `ALTER TABLE` migrations automatically if `history`, `blueprint_name`, or `second_monitor_id` columns are absent from older local databases.*

## 4. Excel Blueprint Architecture (v2 & v2.3 Multi-Sheet Engine)

In v2, blueprints were expanded from simple calendar lists into comprehensive workbooks containing three dedicated sheets parsed dynamically via `openpyxl`. In **v2.3**, `BlueprintEngine` and `BlueprintEditorWindow` expand this to a **4-sheet architecture** by adding native ingestion and export of a master `categories` sheet alongside strict category matching.

### 4.1 Required & Supported Sheet Structure (`.xlsx`)

#### Sheet 1: `championships`

Defines series metadata, prestige, event counts, start types, field parity, parc fermé rules, and tags.

* Columns: `name`, `categories`, `prestige`, `events_count`, `start_type`, `field_type`, `parc_ferme`, `tags` (Writer export order: `Name`, `Categories`, `Prestige`, `Start_Type`, `Field_Type`, `Parc_Ferme`, `Tags`, `Events_Count`).

#### Sheet 2: `events`

Defines track schedules and layout variants mapped to specific championships.

* Columns: `championship`, `track_name`, `variant`.

#### Sheet 3: `cars`

Defines available vehicle models, teams, and category mappings for contract offers, blueprint auto-population, and class filtering.

* Columns: `team_name`, `category`, `car_model` (header detection dynamically resolves columns containing `"team"`, `"cat"`, and `"model"`/`"car"`).

#### Sheet 4: `categories` (Added in v2.3)

Defines the master list of vehicle classes/categories available across the workbook.

* Structure: Column A lists canonical category names (e.g., `GT3`, `GT4`, `LMP1`, `P1`, `F-Ultimate`), automatically skipping header rows (`name`, `category`, `categories`, `class`).

* Parsed dynamically into `engine.categories_list` by detecting any worksheet name containing `"categor"` or matching `"categories"`, feeding dropdown selectors in the In-App Blueprint Editor.

## 5. Intricacies of v2 & v2.3 Core Functionality & Engine Logic

### 5.1 Second Monitor XML Integration & Multiclass Standings Engine

The application interfaces directly with Second Monitor's live XML output (`Championships.xml`) to extract real-world race telemetry and standings.

1. **Namespace Registration & DOM Parsing**:

   * Registers standard XML namespaces globally to handle Second Monitor schema variants (`http://www.w3.org/2001/XMLSchema-instance` and `http://www.w3.org/2001/XMLSchema`).

   * Iterates through `ChampionshipDto` nodes to extract global IDs (`ChampionshipGlobalId`), series names (`ChampionshipName`), and scoring matrices (`Scoring`).

2. **Multiclass & Independent Points Calculation**:

   * Parses driver session results (`DriverSessionResultDto`) containing scratch finish positions (`FinishPosition`), class finish positions (`ClassFinishPosition`), lap data (`WasFastestLap`), and player flags (`IsPlayer`).

   * Because simulation telemetry often bundles multi-class races into single scratch leaderboards, the engine dynamically groups drivers by `ClassName`, sorts them by class finish position, and applies the championship's scoring rules matrix (`scoring_rules`) to compute true independent class standings (`class_standings_map`) and per-round class points gains.

   * When locking in a multiclass round via `lock_in_xml_round_result()`, the engine automatically sets `player_champ_pos` to `player_class_finish` and scales `grid_size` to the class-specific field size (`class_grid_size`) rather than the overall scratch grid.

### 5.2 Persistent Manual Round Result Overrides (`OverrideManager`)

Sim-racers frequently encounter telemetry recording glitches or require manual adjustments to race classifications.

1. **Storage Mechanics**:

   * Overrides are stored as isolated JSON files in `%APPDATA%\AMS2CareerManager\round_overrides\round_overrides_<champ_id>.json`.

   * When a user opens the **Round Editor** (`open_round_editor()`), the app checks for existing overrides. If none exist, it loads raw `RaceResultEntry` dataclass instances from the XML parser.

2. **Serialization Safety**:

   * To prevent `TypeError: Object of type RaceResultEntry is not JSON serializable`, the editor automatically normalizes raw dataclass instances into standard Python dictionaries upon loading. This guarantees seamless JSON dumping when persisting user edits (overall positions, class positions, points gains, fastest laps) without breaking runtime state.

### 5.3 Trajectory Index (TI) & Dynamic Market Evaluation Algorithm (v2 Legacy & v2.3 Advanced Engines)

Season progression and career advancement are governed by the **Trajectory Index (TI)**, a performance metric calculated across completed rounds. In **v2.3**, users can toggle live on the dashboard between **Legacy Ti (Fixed)** and **Advanced Ti (Non-Linear + Recovery)**.

```
               +----------------------------------+
               |     Season Final Evaluation      |
               +----------------+-----------------+
                                |
         +----------------------+----------------------+
         |                      |                      |
         v                      v                      v
+------------------+   +------------------+   +------------------+
| PROMOTION (TI≥75%)|   | RETAINED (40-74%)|   | DEMOTION (TI<40%)|
| - Top tier offers|   | - Lateral/Retain |   | - Fallback drops |
| - Prestige jump  |   | - Same tier pool |   | - Lower series   |
+------------------+   +------------------+   +------------------+

```

1. **Legacy TI Formula (v2)**:
   For each completed race, the driver's championship position ($c$) and race finish position ($f$) are evaluated linearly against the grid size ($g$):
   

   $$
   \text{Score}_{\text{legacy}} = 0.70 \times \left(1.0 - \frac{c - 1}{g - 1}\right) + 0.30 \times \left(1.0 - \frac{f - 1}{g - 1}\right)
   $$

2. **Advanced Non-Linear TI & Field-Percentage Recovery Formula (v2.3 Default)**:
   Evaluates championship position ($c$), race finish position ($f$), and qualifying position ($q$) against grid size ($g \ge 2$) using non-linear power curves that reward podium/front-running excellence, combined with a graduated field-percentage recovery delta:
   

   $$
   \text{Score}_{\text{champ}} = \left(1.0 - \frac{c - 1}{g - 1}\right)^{1.25}, \quad \text{Score}_{\text{race}} = \left(1.0 - \frac{f - 1}{g - 1}\right)^{1.25}, \quad \text{Score}_{\text{quali}} = \left(1.0 - \frac{q - 1}{g - 1}\right)^{1.1}
   $$

   
   Let $\Delta_{\text{field}} = \frac{q - f}{g - 1}$ represent the fraction of the field gained or lost from qualifying to race finish. The graduated recovery bonus/penalty is computed as:
   

   $$
   B_{\text{recovery}} = \begin{cases} \min\left(0.35,\; \left(\Delta_{\text{field}}\right)^{0.8} \times 0.40\right) & \text{if } \Delta_{\text{field}} > 0 \\ \max\left(-0.15,\; \Delta_{\text{field}} \times 0.15\right) & \text{if } \Delta_{\text{field}} < 0 \\ 0.0 & \text{if } \Delta_{\text{field}} = 0 \end{cases}
   $$

   
   The weighted weekend score (clamped to $[0.0, 1.15]$ per round) and season Trajectory Index ($TI \in [0.0, 1.0]$) are:
   

   $$
   \text{Score}_{\text{weekend}} = \text{clamp}\Big(0.40 \times \text{Score}_{\text{champ}} + 0.35 \times \text{Score}_{\text{race}} + 0.25 \times \text{Score}_{\text{quali}} + B_{\text{recovery}},\; 0.0,\; 1.15\Big)
   $$

   $$
   TI = \text{clamp}\left(\frac{1}{N} \sum_{i=1}^{N} \text{Score}_{\text{weekend}, i},\; 0.0,\; 1.0\right)
   $$

3. **Contract Market Continuum**:

   * **TI** $\ge$ **0.75 (Promotion)**: Unlocks retention offers, higher-tier promotion seats (up to $+5 + \lfloor \frac{TI - 0.75}{0.25} \times 10 \rfloor$ prestige), lateral moves, and a top-of-lower-tier fallback option.

   * **0.40** $\le$ **TI** $<$ **0.75 (Retention/Lateral)**: Offers current tier retention (1–2 seats), lateral series swaps, and up to 2 lower-tier fallback options.

   * **TI** $<$ **0.40 (Demotion)**: Generates fallback and demotion offers into lower-prestige series scaled by demotion drop depth ($5 + \lfloor (1.0 - TI) \times 20 \rfloor$).

   * **Strict Category Matching (v2.3 Enhancement)**: Across contract market generation (`get_strictly_matched_cars`), initial career setup (`create`), and blueprint auto-population (`auto_populate_teams_for_championship`), category comparison enforces exact case-insensitive string equality (`cc == car_cat`) rather than substring checks, eliminating shorthand category leaks (e.g., preventing `"P1"` from erroneously matching `"LMP1"`).

### 5.4 In-App Blueprint Creator & Editor (`BlueprintEditorWindow` - v2.3)

Version 2.3 eliminates the need for external spreadsheet software to author or modify career blueprints:

1. **Staged Three-Tab Authoring Environment**: Loads existing active blueprint data into memory (`staged_championships`, `staged_events`, `staged_cars`) or starts from a clean slate across **Championships**, **Events & Tracks**, and **Cars & Teams** tabs.

2. **Inline Championship & Calendar Builder**: Double-clicking any championship (or clicking *Edit Selected Championship*) opens `open_championship_form()`, allowing users to edit series names, prestige ($1\text{–}100$), multi-class category lists (fed from `categories` sheet Column A + staged categories), round counts, start types (`Rolling`/`Standing`), and custom round-by-round track/variant calendars.

3. **Automated 40-Car Real-Team Grid Population**: When creating a new championship (`auto_populate_teams_for_championship()`), the editor harvests authentic real team names from the workbook's existing car pool (filtering out generic placeholders), strictly matches valid car models per assigned category, and generates 40 unique team-and-car pairings distributed across the championship's classes.

4. **4-Sheet Workbook Export**: `save_as_new_blueprint()` writes all staged data cleanly into `championships`, `events`, `cars`, and `categories` worksheets inside `%APPDATA%\AMS2CareerManager\blueprints\` and immediately refreshes the main screen's blueprint library.

### 5.5 Save Blueprint Re-Assignment & Content Summary Viewer (v2.3)

* Accessed directly from the Dashboard metadata strip via **Update Save Blueprint** (`change_save_blueprint()`), users can switch or import a new `.xlsx` blueprint file for an existing career save without losing season progress or race history.

* Upon confirmation, the application updates `blueprint_name` in SQLite, reloads `engine`, refreshes the dashboard, and launches a **Blueprint Re-bound & Content Summary** modal displaying all loaded championships, prestige ratings, car categories, and event counts.

## 6. UI/UX Specifications (v2 & v2.3 Four-Screen Architecture + Modal Blueprint Editor)

### 6.1 Palette & Ergonomics

* Built upon the slate-dark palette (`#090d16` primary background, `#121824` / `#1e293b` surface panels, `#2563eb` / `#3b82f6` primary accent, `#059669` success/confirm, `#d97706` warning/edit, `#dc2626` danger/delete, `#f8fafc` high-contrast text).

* Features zebra-striped `ttk.Treeview` tables (`#1e293b` / `#162032`), dynamic comboboxes, header version badge (`v2.3`), and safety confirmation modals.

### 6.2 Screen & Modal Architecture

1. **Screen 1: Career Select & Blueprint Library**: Manages profile creation, JSON backup export/import, profile deletion, live `.xlsx` blueprint library selection, folder access, and launching the **Blueprint Editor**.

2. **Screen 2: Unified Dashboard & Tabbed Inspector**:

   * *Top Metadata Strip*: Displays series regulations (`Start`, `Field`, `Parc Fermé`, `Prestige`, `Classes`, bound `Blueprint`), the live **Ti System** mode selector (`Legacy Ti (Fixed)` vs. `Advanced Ti (Non-Linear + Recovery)`), and the **Update Save Blueprint** button.

   * *Left Column*: Rank Status indicator, Season Event Calendar (with interactive track/layout editing), and Championship Classes & Eligible Cars breakdown.

   * *Right Column (Tabbed)*: Second Monitor XML link manager, Live XML Championship Standings (with class scope filtering), and Per-Round Classification & Result Recorder (featuring round class scope filtering, auto-detected qualifying positions, grid sizing, and the **Round XML Editor** modal).

3. **Screen 3: Summary & Market Offers**: End-of-season evaluation screen displaying tiered contract offers with strict car class and team pairings, enabling instant contract signing and history archiving.

4. **Screen 4: Global History & Save Inspector**: Multi-season archive viewer allowing inspection of past career profiles, completed season scores, and historical race log breakdowns.

5. **Modal Suite: In-App Blueprint Creator & Editor (`BlueprintEditorWindow`)**: Full-featured staging editor with inline championship/calendar editing, 40-car auto-population, and `.xlsx` export.

## 7. Complete v2.3 Consolidated Codebase Implementation

```
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import sqlite3
import os
import random
import openpyxl
import ctypes
import sys
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional

try:
    # Forces Windows taskbar to treat this as its own unique application
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("AMS2CareerManager.Desktop.App")
except Exception:
    pass

# SECURE APPDATA PERSISTENCE SETUP
APPDATA_DIR = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'AMS2CareerManager')
os.makedirs(APPDATA_DIR, exist_ok=True)
DB_FILE = os.path.join(APPDATA_DIR, "ams2_careers.db")

# BLUEPRINTS DIRECTORY SETUP
BLUEPRINTS_DIR = os.path.join(APPDATA_DIR, "blueprints")
os.makedirs(BLUEPRINTS_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS careers (
            id TEXT PRIMARY KEY,
            name TEXT,
            driver_name TEXT,
            season INTEGER,
            current_contract TEXT,
            races TEXT,
            calendar TEXT,
            history TEXT,
            blueprint_name TEXT DEFAULT '',
            second_monitor_id TEXT DEFAULT ''
        )
    """)
    c.execute("PRAGMA table_info(careers)")
    columns = [col[1] for col in c.fetchall()]
    if "history" not in columns:
        try:
            c.execute("ALTER TABLE careers ADD COLUMN history TEXT DEFAULT '[]'")
        except Exception:
            pass
    if "blueprint_name" not in columns:
        try:
            c.execute("ALTER TABLE careers ADD COLUMN blueprint_name TEXT DEFAULT ''")
        except Exception:
            pass
    if "second_monitor_id" not in columns:
        try:
            c.execute("ALTER TABLE careers ADD COLUMN second_monitor_id TEXT DEFAULT ''")
        except Exception:
            pass
    conn.commit()
    conn.close()

init_db()

# Register standard XML namespaces globally for Second Monitor compatibility
ET.register_namespace('', "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace('xsd', "http://www.w3.org/2001/XMLSchema")


# --- ADVANCED PARSER DATA STRUCTURES & PARSER ---

@dataclass
class DriverStanding:
    driver_id: str
    driver_name: str
    position: int
    total_points: float
    class_id: Optional[str] = None
    class_name: Optional[str] = None

@dataclass
class RaceResultEntry:
    driver_name: str
    car_name: str
    class_id: str
    class_name: str
    finish_position: int
    class_finish_position: int
    total_points: float
    points_gain: float
    is_player: bool = False
    fastest_lap: bool = False

@dataclass
class SessionRaceResults:
    session_id: str
    session_name: str
    results: List[RaceResultEntry] = field(default_factory=list)

@dataclass
class ChampionshipData:
    championship_name: str
    championship_id: str
    overall_standings: List[DriverStanding] = field(default_factory=list)
    class_standings: Dict[str, List[DriverStanding]] = field(default_factory=dict)
    session_results: List[SessionRaceResults] = field(default_factory=list)
    rounds: List[dict] = field(default_factory=list)
    scoring_rules: List[float] = field(default_factory=list)


class SecondMonitorAdvancedParser:
    """Advanced parser for Second Monitor Championships.xml files.
    Fully supports robust multiclass extraction, scoring rules parsing, and independent class standings.
    """
    def __init__(self, xml_path=None):
        if xml_path is None:
            self.xml_path = Path.home() / "Documents" / "SecondMonitor" / "Championships" / "Championships.xml"
        else:
            self.xml_path = Path(xml_path)

    def load_tree(self):
        if not self.xml_path.exists():
            return None, None
        try:
            tree = ET.parse(self.xml_path)
            root = tree.getroot()
            return tree, root
        except Exception as e:
            print(f"Error loading XML: {e}")
            return None, None

    def parse_championships(self) -> Dict[str, ChampionshipData]:
        tree, root = self.load_tree()
        if root is None:
            return {}

        championships_data = {}
        
        for champ_elem in root.findall(".//ChampionshipDto"):
            champ_id = champ_elem.get("ChampionshipGlobalId") or champ_elem.findtext("ChampionshipGlobalId", "")
            champ_name = champ_elem.get("ChampionshipName") or champ_elem.findtext("ChampionshipName", "Unknown Championship")
            
            # Extract Scoring Rules Matrix
            scoring_rules = []
            scoring_node = champ_elem.find(".//Scoring")
            if scoring_node is not None:
                for pts_elem in scoring_node.findall(".//*"):
                    val_text = pts_elem.text
                    if val_text:
                        try:
                            scoring_rules.append(float(val_text.strip()))
                        except ValueError:
                            pass
            if not scoring_rules:
                # Default fallback scoring matrix if XML scoring node format varies
                scoring_rules = [350, 320, 300, 280, 260, 250, 240, 230, 220, 210, 200, 190, 180, 170, 160, 150, 140, 130, 120, 110, 100, 90, 80, 70, 60, 50, 40, 30, 20, 10]

            # Extract Overall Standings
            overall_standings = []
            drivers_node = champ_elem.find("Drivers")
            if drivers_node is not None:
                for driver_elem in drivers_node.findall("DriverDto"):
                    overall_standings.append(DriverStanding(
                        driver_id=driver_elem.findtext("DriverId", "") or driver_elem.get("DriverId", ""),
                        driver_name=driver_elem.get("LastUsedName") or driver_elem.findtext("LastUsedName", "Unknown"),
                        position=int(driver_elem.get("Position") or driver_elem.findtext("Position", "0")),
                        total_points=float(driver_elem.get("TotalPoints") or driver_elem.findtext("TotalPoints", "0.0"))
                    ))
            overall_standings.sort(key=lambda x: x.position)

            # Extract Rounds & Session Results first so we can aggregate independent class points accurately
            rounds = []
            events_node = champ_elem.find("Events")
            events = events_node.findall("EventDto") if events_node is not None else champ_elem.findall(".//EventDto")

            session_results_list = []
            for r_idx, event in enumerate(events, start=1):
                track_name = event.get("TrackName") or event.findtext("TrackName", f"Round {r_idx}")
                event_status = event.get("EventStatus") or event.findtext("EventStatus", "")
                
                sessions = []
                sessions_node = event.find("Sessions")
                session_list = sessions_node.findall("SessionDto") if sessions_node is not None else event.findall(".//SessionDto")

                for session in session_list:
                    session_type = session.get("SessionType") or session.findtext("SessionType", "Race")
                    session_id = session.findtext("SessionId", "")
                    session_name = session.findtext("SessionName", session_type)
                    
                    results = []
                    res_nodes = session.findall(".//DriverSessionResultDto")
                    for res in res_nodes:
                        cls_name = res.get("ClassName") or res.findtext("ClassName", "General")
                        results.append(RaceResultEntry(
                            driver_name=res.get("DriverName") or res.findtext("DriverName", "Unknown"),
                            car_name=res.get("CarName") or res.findtext("CarName", ""),
                            class_id=res.get("ClassId") or res.findtext("ClassId", ""),
                            class_name=cls_name,
                            finish_position=int(res.get("FinishPosition") or res.findtext("FinishPosition", "0")),
                            class_finish_position=int(res.get("ClassFinishPosition") or res.findtext("ClassFinishPosition", "0")),
                            total_points=float(res.get("TotalPoints") or res.findtext("TotalPoints", "0.0")),
                            points_gain=float(res.get("PointsGain") or res.findtext("PointsGain", "0.0")),
                            is_player=(res.get("IsPlayer") or res.findtext("IsPlayer", "false")).lower() == "true",
                            fastest_lap=(res.get("WasFastestLap") or res.findtext("WasFastestLap", "false")).lower() == "true"
                        ))
                    results.sort(key=lambda x: x.finish_position)
                    
                    sessions.append({
                        "session_type": session_type,
                        "session_name": session_name,
                        "results": results
                    })
                    
                    session_results_list.append(SessionRaceResults(
                        session_id=session_id,
                        session_name=session_name,
                        results=results
                    ))
                
                rounds.append({
                    "round_number": r_idx,
                    "track_name": track_name,
                    "event_status": event_status,
                    "sessions": sessions
                })

            # Dynamically Calculate Independent Class Standings using XML Scoring Rules Matrix
            class_standings_map = {}
            class_driver_totals = {} # {class_name: {driver_name: total_points}}
            class_driver_positions = {} # {class_name: {driver_name: last_pos}}

            for round_data in rounds:
                for sess in round_data["sessions"]:
                    if "race" in sess["session_type"].lower():
                        # Group race results by class
                        class_groups = {}
                        for r in sess["results"]:
                            cn = r.class_name if r.class_name else "General"
                            if cn not in class_groups:
                                class_groups[cn] = []
                            class_groups[cn].append(r)
                        
                        for cn, group_results in class_groups.items():
                            # Sort group by finish position or class finish position
                            group_results.sort(key=lambda x: x.class_finish_position if x.class_finish_position > 0 else x.finish_position)
                            
                            if cn not in class_driver_totals:
                                class_driver_totals[cn] = {}
                            if cn not in class_driver_positions:
                                class_driver_positions[cn] = {}

                            for idx, r in enumerate(group_results, start=1):
                                d_name = r.driver_name
                                # Assign points based on scoring rules index (idx - 1)
                                pts = scoring_rules[idx - 1] if (idx - 1) < len(scoring_rules) else max(5.0, scoring_rules[-1] - ((idx - len(scoring_rules)) * 5))
                                
                                class_driver_totals[cn][d_name] = class_driver_totals[cn].get(d_name, 0.0) + pts
                                class_driver_positions[cn][d_name] = idx

            for cn, totals in class_driver_totals.items():
                sorted_drivers = sorted(totals.items(), key=lambda x: x[1], reverse=True)
                class_standings_list = []
                for rank, (d_name, pts) in enumerate(sorted_drivers, start=1):
                    class_standings_list.append(DriverStanding(
                        driver_id="",
                        driver_name=d_name,
                        position=rank,
                        total_points=pts,
                        class_name=cn
                    ))
                class_standings_map[cn] = class_standings_list

            championships_data[champ_id] = ChampionshipData(
                championship_name=champ_name,
                championship_id=champ_id,
                overall_standings=overall_standings,
                class_standings=class_standings_map,
                session_results=session_results_list,
                rounds=rounds,
                scoring_rules=scoring_rules
            )
            
        return championships_data


class OverrideManager:
    """Manages persistent manual overrides for round results in JSON format."""
    def __init__(self, storage_dir=APPDATA_DIR):
        self.storage_dir = Path(storage_dir) / "round_overrides"
        self.storage_dir.mkdir(exist_ok=True)

    def get_override_path(self, champ_id):
        safe_id = "".join(c if c.isalnum() else "_" for c in champ_id)
        return self.storage_dir / f"round_overrides_{safe_id}.json"

    def load_overrides(self, champ_id):
        path = self.get_override_path(champ_id)
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_overrides(self, champ_id, overrides):
        path = self.get_override_path(champ_id)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(overrides, f, indent=4)


class BlueprintEngine:
    def __init__(self):
        self.championships = []
        self.events = []
        self.cars = []
        self.categories_list = []  # Dynamic Master Categories list parsed from Column A of category sheets
        self.is_loaded = False
        self.active_blueprint_filename = ""

    def load_from_excel(self, file_path):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        
        # Dynamically parse master categories from any sheet named "categories" or containing "categor"
        self.categories_list = []
        cat_sheet_name = next((s for s in wb.sheetnames if "categor" in s.lower() or s.lower() == "categories"), None)
        if cat_sheet_name:
            sheet = wb[cat_sheet_name]
            for row in sheet.iter_rows(min_row=1, values_only=True):
                if row and row[0] is not None:
                    val = str(row[0]).strip()
                    # Skip header if it equals "name", "category", etc.
                    if val and val.lower() not in ["name", "category", "categories", "class"]:
                        if val not in self.categories_list:
                            self.categories_list.append(val)

        if "championships" in wb.sheetnames:
            sheet = wb["championships"]
            headers = [str(cell.value).strip().lower() for cell in sheet[1] if cell.value is not None]
            name_idx = headers.index("name") if "name" in headers else 0
            cat_idx = headers.index("categories") if "categories" in headers else 1
            prest_idx = headers.index("prestige") if "prestige" in headers else 2
            events_idx = headers.index("events_count") if "events_count" in headers else 8
            
            start_idx = next((i for i, h in enumerate(headers) if "start" in h), None)
            field_idx = next((i for i, h in enumerate(headers) if "field" in h), None)
            parc_idx = next((i for i, h in enumerate(headers) if "parc" in h or "ferme" in h), None)
            tags_idx = next((i for i, h in enumerate(headers) if "tag" in h), None)
            
            self.championships = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[name_idx]: continue
                categories_raw = str(row[cat_idx]) if len(row) > cat_idx and row[cat_idx] else ""
                categories = [c.strip() for c in categories_raw.split(",") if c.strip()]
                
                start_type = str(row[start_idx]).strip() if start_idx is not None and len(row) > start_idx and row[start_idx] else "Rolling"
                field_type_val = str(row[field_idx]).strip() if field_idx is not None and len(row) > field_idx and row[field_idx] else ""
                field_type = "Identical" if field_type_val and field_type_val.lower() != "none" else "Non-Identical"
                
                parc_raw = str(row[parc_idx]).strip().lower() if parc_idx is not None and len(row) > parc_idx and row[parc_idx] is not None else "false"
                parc_ferme = "True" if parc_raw in ["true", "1", "yes"] else "False"

                tags_raw = str(row[tags_idx]).strip() if tags_idx is not None and len(row) > tags_idx and row[tags_idx] else ""
                tags = [t.strip() for t in tags_raw.split(",") if t.strip()]

                self.championships.append({
                    "name": str(row[name_idx]).strip(),
                    "categories": categories,
                    "tags": tags,
                    "prestige": int(row[prest_idx] or 50),
                    "events_count": int(row[events_idx] or 8),
                    "start_type": start_type,
                    "field_type": field_type,
                    "parc_ferme": parc_ferme
                })

        if "events" in wb.sheetnames:
            sheet = wb["events"]
            self.events = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                champ_name = str(row[0]).strip()
                track_name = str(row[1]).strip()
                variant = str(row[2]).strip() if len(row) > 2 and row[2] else "Grand Prix"
                self.events.append({
                    "championship": champ_name,
                    "track_name": track_name,
                    "variant": variant
                })

        if "cars" in wb.sheetnames:
            sheet = wb["cars"]
            headers = [str(cell.value).strip().lower() for cell in sheet[1] if cell.value is not None]
            
            t_idx = next((i for i, h in enumerate(headers) if "team" in h), 0)
            c_idx = next((i for i, h in enumerate(headers) if "cat" in h), 1)
            m_idx = next((i for i, h in enumerate(headers) if "model" in h or "car" in h), 2 if len(headers) > 2 else 0)

            self.cars = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                team_name = str(row[t_idx]).strip() if len(row) > t_idx and row[t_idx] else ""
                category = str(row[c_idx]).strip() if len(row) > c_idx and row[c_idx] else "General"
                car_model = str(row[m_idx]).strip() if len(row) > m_idx and row[m_idx] else "Spec Race Car"
                
                if team_name:
                    self.cars.append({
                        "team_name": team_name,
                        "category": category,
                        "car_model": car_model
                    })

        self.is_loaded = True
        self.active_blueprint_filename = os.path.basename(file_path)
        return len(self.championships)

    def generate_calendar_for_champ(self, champ_name):
        champ = next((c for c in self.championships if c['name'].lower() == champ_name.lower()), None)
        target_count = champ['events_count'] if champ else 6
        matching_events = [e for e in self.events if e['championship'].lower() == champ_name.lower()]
        
        calendar = []
        if matching_events:
            for i in range(target_count):
                ev = matching_events[i % len(matching_events)]
                calendar.append(f"{ev['track_name']} ({ev['variant']})")
        else:
            calendar = [f"Standard Circuit R{i+1}" for i in range(target_count)]
            
        return calendar

engine = BlueprintEngine()


# --- IN-APP BLUEPRINT CREATOR & EDITOR WINDOW ---

class BlueprintEditorWindow(tk.Toplevel):
    """Interactive In-App Blueprint Creator and Editor.
    Allows users to view, add, modify, and edit Championships, Events, and Cars,
    featuring dynamic dropdowns fed from parsed workbook data and inline editing.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.title("In-App Blueprint Creator & Editor")
        self.geometry("980x680")
        self.configure(bg="#090d16")
        self.transient(parent)
        self.grab_set()

        # Local staged data copies
        self.staged_championships = [dict(c) for c in engine.championships] if engine.is_loaded else []
        self.staged_events = [dict(e) for e in engine.events] if engine.is_loaded else []
        self.staged_cars = [dict(car) for car in engine.cars] if engine.is_loaded else []

        self.create_editor_widgets()

    def get_all_available_categories(self) -> List[str]:
        """Extracts all unique car categories/classes from the master categories sheet column A, staged cars, and existing championships."""
        cats = set()
        # 1. Pull from parsed engine categories list (Column A of category sheet)
        for c in engine.categories_list:
            if c:
                cats.add(c.strip())
        # 2. Pull from staged cars
        for car in self.staged_cars:
            c = car.get('category', '').strip()
            if c:
                cats.add(c)
        # 3. Pull from staged championships
        for champ in self.staged_championships:
            for cat in champ.get('categories', []):
                c = cat.strip()
                if c:
                    cats.add(c)
        return sorted(list(cats))

    def get_all_available_tracks(self) -> List[str]:
        """Extracts all unique track names from staged events."""
        tracks = set()
        for ev in self.staged_events:
            t = ev.get('track_name', '').strip()
            if t:
                tracks.add(t)
        if not tracks:
            tracks = {"Interlagos", "Monza", "Spa-Francorchamps", "Silverstone", "Nürburgring", "Daytona", "Bathurst"}
        return sorted(list(tracks))

    def create_editor_widgets(self):
        header_lbl = tk.Label(self, text="Blueprint Staging Editor & Creator", font=("Helvetica", 13, "bold"), bg="#090d16", fg="#f8fafc")
        header_lbl.pack(padx=15, pady=(15, 5), anchor="w")

        info_lbl = tk.Label(self, text="Double-click any championship to edit its details, tracks, and car categories inline, or add/delete items below.", font=("Helvetica", 8, "italic"), bg="#090d16", fg="#38bdf8")
        info_lbl.pack(padx=15, pady=(0, 10), anchor="w")

        # Notebook for Categories
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        # --- TAB 1: CHAMPIONSHIPS ---
        self.tab_champs = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_champs, text="Championships")

        champ_frame = tk.Frame(self.tab_champs, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        champ_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        champ_cols = ("name", "prestige", "categories", "events_count", "start_type", "field_type", "parc_ferme")
        self.champ_tree = ttk.Treeview(champ_frame, columns=champ_cols, show="headings", height=12)
        self.champ_tree.heading("name", text="Championship Name")
        self.champ_tree.heading("prestige", text="Prestige")
        self.champ_tree.heading("categories", text="Categories / Classes")
        self.champ_tree.heading("events_count", text="Rounds")
        self.champ_tree.heading("start_type", text="Start")
        self.champ_tree.heading("field_type", text="Field")
        self.champ_tree.heading("parc_ferme", text="Parc Fermé")

        self.champ_tree.column("name", width=220)
        self.champ_tree.column("prestige", width=65, anchor="center")
        self.champ_tree.column("categories", width=180)
        self.champ_tree.column("events_count", width=60, anchor="center")
        self.champ_tree.column("start_type", width=80, anchor="center")
        self.champ_tree.column("field_type", width=90, anchor="center")
        self.champ_tree.column("parc_ferme", width=80, anchor="center")

        self.champ_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.champ_tree.bind("<Double-1>", self.on_championship_double_click)

        c_scroll = ttk.Scrollbar(champ_frame, orient=tk.VERTICAL, command=self.champ_tree.yview)
        self.champ_tree.configure(yscrollcommand=c_scroll.set)
        c_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        champ_btn_bar = tk.Frame(self.tab_champs, bg="#090d16", pady=5)
        champ_btn_bar.pack(fill=tk.X, padx=10)
        tk.Button(champ_btn_bar, text="+ Add Championship", command=self.add_championship_dialog, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(champ_btn_bar, text="Edit Selected Championship", command=self.edit_selected_championship, bg="#d97706", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(champ_btn_bar, text="Delete Selected Championship", command=self.delete_championship, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        # --- TAB 2: EVENTS & CALENDAR ---
        self.tab_events = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_events, text="Events & Tracks")

        event_frame = tk.Frame(self.tab_events, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        event_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        event_cols = ("championship", "track_name", "variant")
        self.event_tree = ttk.Treeview(event_frame, columns=event_cols, show="headings", height=12)
        self.event_tree.heading("championship", text="Championship Series")
        self.event_tree.heading("track_name", text="Track Name")
        self.event_tree.heading("variant", text="Layout Variant")

        self.event_tree.column("championship", width=300)
        self.event_tree.column("track_name", width=250)
        self.event_tree.column("variant", width=180)

        self.event_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        e_scroll = ttk.Scrollbar(event_frame, orient=tk.VERTICAL, command=self.event_tree.yview)
        self.event_tree.configure(yscrollcommand=e_scroll.set)
        e_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        event_btn_bar = tk.Frame(self.tab_events, bg="#090d16", pady=5)
        event_btn_bar.pack(fill=tk.X, padx=10)
        tk.Button(event_btn_bar, text="+ Add Event / Track", command=self.add_event_dialog, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(event_btn_bar, text="Delete Selected Event", command=self.delete_event, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        # --- TAB 3: CARS & TEAMS ---
        self.tab_cars = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_cars, text="Cars & Teams")

        car_frame = tk.Frame(self.tab_cars, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        car_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        car_cols = ("team_name", "category", "car_model")
        self.car_tree = ttk.Treeview(car_frame, columns=car_cols, show="headings", height=12)
        self.car_tree.heading("team_name", text="Team Name")
        self.car_tree.heading("category", text="Car Class Category")
        self.car_tree.heading("car_model", text="Car Model / Vehicle")

        self.car_tree.column("team_name", width=260)
        self.car_tree.column("category", width=200)
        self.car_tree.column("car_model", width=280)

        self.car_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        car_scroll = ttk.Scrollbar(car_frame, orient=tk.VERTICAL, command=self.car_tree.yview)
        self.car_tree.configure(yscrollcommand=car_scroll.set)
        car_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        car_btn_bar = tk.Frame(self.tab_cars, bg="#090d16", pady=5)
        car_btn_bar.pack(fill=tk.X, padx=10)
        tk.Button(car_btn_bar, text="+ Add Car Model", command=self.add_car_dialog, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(car_btn_bar, text="Delete Selected Car", command=self.delete_car, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        # Bottom Action Bar
        bottom_bar = tk.Frame(self, bg="#090d16", pady=10)
        bottom_bar.pack(fill=tk.X, padx=15)
        tk.Button(bottom_bar, text="Save as New Blueprint (.xlsx)", command=self.save_as_new_blueprint, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), padx=20, pady=8, relief=tk.FLAT).pack(side=tk.RIGHT)

        self.refresh_editor_views()

    def refresh_editor_views(self):
        for i in self.champ_tree.get_children(): self.champ_tree.delete(i)
        for idx, c in enumerate(self.staged_championships):
            cats = ", ".join(c.get('categories', []))
            self.champ_tree.insert("", tk.END, iid=str(idx), values=(c['name'], c['prestige'], cats, c['events_count'], c['start_type'], c['field_type'], c['parc_ferme']))

        for i in self.event_tree.get_children(): self.event_tree.delete(i)
        for idx, e in enumerate(self.staged_events):
            self.event_tree.insert("", tk.END, iid=str(idx), values=(e['championship'], e['track_name'], e['variant']))

        for i in self.car_tree.get_children(): self.car_tree.delete(i)
        for idx, car in enumerate(self.staged_cars):
            self.car_tree.insert("", tk.END, iid=str(idx), values=(car['team_name'], car['category'], car['car_model']))

    def add_championship_dialog(self):
        self.open_championship_form(mode="add")

    def edit_selected_championship(self):
        sel = self.champ_tree.selection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a championship to edit.")
            return
        idx = int(sel[0])
        self.open_championship_form(mode="edit", index=idx)

    def on_championship_double_click(self, event):
        sel = self.champ_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        self.open_championship_form(mode="edit", index=idx)

    def open_championship_form(self, mode="add", index=None):
        dlg = tk.Toplevel(self)
        dlg.title("Edit Championship" if mode == "edit" else "Add New Championship")
        dlg.geometry("540x620")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        champ_data = self.staged_championships[index] if mode == "edit" and index is not None else {
            "name": "", "prestige": 50, "categories": [], "events_count": 8, "start_type": "Rolling", "field_type": "Non-Identical", "parc_ferme": "False"
        }

        tk.Label(dlg, text="Championship Name:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(12, 2))
        e_name = ttk.Entry(dlg, width=55)
        e_name.pack(padx=20)
        e_name.insert(0, champ_data["name"])

        tk.Label(dlg, text="Prestige (1-100):", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(8, 2))
        e_prest = ttk.Spinbox(dlg, from_=1, to=100, width=12)
        e_prest.pack(anchor="w", padx=20)
        e_prest.set(champ_data["prestige"])

        tk.Label(dlg, text="Car Category / Class (Select from workbook or type a brand new class):", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(8, 2))
        
        cat_frame = tk.Frame(dlg, bg="#090d16")
        cat_frame.pack(fill=tk.X, padx=20)

        available_cats = self.get_all_available_categories()
        if not available_cats:
            available_cats = ["GT3", "GT4", "F-3", "F-Ultimate", "Prototype", "Touring Cars"]

        cat_combo = ttk.Combobox(cat_frame, values=available_cats, width=40)
        cat_combo.pack(side=tk.LEFT, padx=(0, 8))
        existing_cats = champ_data.get("categories", [])
        if existing_cats:
            cat_combo.set(existing_cats[0])
        else:
            cat_combo.set(available_cats[0])

        selected_cats_list = list(existing_cats)

        cat_listbox_frame = tk.Frame(dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        cat_listbox_frame.pack(fill=tk.BOTH, expand=False, padx=20, pady=(4, 6))

        cat_listbox = tk.Listbox(cat_listbox_frame, font=("Helvetica", 8), bg="#090d16", fg="#f8fafc", height=3)
        cat_listbox.pack(fill=tk.BOTH, expand=True)

        for c in selected_cats_list:
            cat_listbox.insert(tk.END, c)

        def add_category_to_list():
            val = cat_combo.get().strip()
            if val and val not in selected_cats_list:
                selected_cats_list.append(val)
                cat_listbox.insert(tk.END, val)

        def remove_category_from_list():
            sel = cat_listbox.curselection()
            if sel:
                idx = sel[0]
                val = cat_listbox.get(idx)
                cat_listbox.delete(idx)
                if val in selected_cats_list:
                    selected_cats_list.remove(val)

        cat_btn_row = tk.Frame(dlg, bg="#090d16")
        cat_btn_row.pack(fill=tk.X, padx=20, pady=(0, 6))
        tk.Button(cat_btn_row, text="+ Add Class", command=add_category_to_list, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(cat_btn_row, text="Remove Selected", command=remove_category_from_list, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT)

        tk.Label(dlg, text="Events Count (Rounds):", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(6, 2))
        e_rounds = ttk.Spinbox(dlg, from_=1, to=30, width=12)
        e_rounds.pack(anchor="w", padx=20)
        e_rounds.set(champ_data["events_count"])

        tk.Label(dlg, text="Start Type:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(6, 2))
        e_start = ttk.Combobox(dlg, values=["Rolling", "Standing"], state="readonly", width=18)
        e_start.pack(anchor="w", padx=20)
        e_start.set(champ_data["start_type"])

        # --- TRACKS BUILDER FOR THIS CHAMPIONSHIP ---
        tk.Label(dlg, text="Championship Calendar / Tracks Builder:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#38bdf8").pack(anchor="w", padx=20, pady=(8, 2))
        
        track_row = tk.Frame(dlg, bg="#090d16")
        track_row.pack(fill=tk.X, padx=20)

        available_tracks = self.get_all_available_tracks()
        track_combo = ttk.Combobox(track_row, values=available_tracks, width=32)
        track_combo.pack(side=tk.LEFT, padx=(0, 8))
        if available_tracks:
            track_combo.set(available_tracks[0])

        variant_entry = ttk.Entry(track_row, width=15)
        variant_entry.pack(side=tk.LEFT, padx=(0, 8))
        variant_entry.insert(0, "Grand Prix")

        # Load existing events for this championship if editing
        champ_name_current = champ_data["name"]
        champ_events = [e for e in self.staged_events if e['championship'].lower() == champ_name_current.lower()] if mode == "edit" else []

        track_listbox_frame = tk.Frame(dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        track_listbox_frame.pack(fill=tk.BOTH, expand=False, padx=20, pady=(4, 6))

        track_listbox = tk.Listbox(track_listbox_frame, font=("Helvetica", 8), bg="#090d16", fg="#f8fafc", height=4)
        track_listbox.pack(fill=tk.BOTH, expand=True)

        staged_champ_events_list = list(champ_events)
        for ev in staged_champ_events_list:
            track_listbox.insert(tk.END, f"{ev['track_name']} ({ev['variant']})")

        def add_track_to_champ():
            t_name = track_combo.get().strip()
            v_name = variant_entry.get().strip() or "Grand Prix"
            if t_name:
                item_str = f"{t_name} ({v_name})"
                track_listbox.insert(tk.END, item_str)
                staged_champ_events_list.append({
                    "championship": e_name.get().strip() or "New Championship",
                    "track_name": t_name,
                    "variant": v_name
                })

        def remove_track_from_champ():
            sel = track_listbox.curselection()
            if sel:
                idx = sel[0]
                track_listbox.delete(idx)
                if idx < len(staged_champ_events_list):
                    staged_champ_events_list.pop(idx)

        track_btn_row = tk.Frame(dlg, bg="#090d16")
        track_btn_row.pack(fill=tk.X, padx=20, pady=(0, 10))
        tk.Button(track_btn_row, text="+ Add Track", command=add_track_to_champ, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(track_btn_row, text="Remove Track", command=remove_track_from_champ, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT)

        def save_championship_form():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Error", "Championship name cannot be empty.")
                return
            if not selected_cats_list:
                selected_cats_list.append("General")

            rounds_count = int(e_rounds.get())

            updated_entry = {
                "name": name,
                "categories": selected_cats_list,
                "tags": champ_data.get("tags", ["Custom"]),
                "prestige": int(e_prest.get()),
                "events_count": rounds_count,
                "start_type": e_start.get(),
                "field_type": champ_data.get("field_type", "Non-Identical"),
                "parc_ferme": champ_data.get("parc_ferme", "False")
            }

            if mode == "edit" and index is not None:
                old_name = self.staged_championships[index]['name']
                self.staged_championships[index] = updated_entry
                
                # Update events associated with this championship name
                self.staged_events = [e for e in self.staged_events if e['championship'].lower() != old_name.lower()]
                for ev in staged_champ_events_list:
                    ev['championship'] = name
                    self.staged_events.append(ev)
                added_count = 0
            else:
                self.staged_championships.append(updated_entry)
                for ev in staged_champ_events_list:
                    ev['championship'] = name
                    self.staged_events.append(ev)

                # If no tracks were explicitly added, generate default placeholder events
                if not staged_champ_events_list:
                    default_tracks = ["Interlagos", "Monza", "Spa-Francorchamps", "Silverstone", "Nürburgring", "Daytona", "Bathurst", "Brands Hatch"]
                    for i in range(rounds_count):
                        t_name = default_tracks[i % len(default_tracks)]
                        self.staged_events.append({
                            "championship": name,
                            "track_name": t_name,
                            "variant": "Grand Prix"
                        })

                # Auto-populate 40 authentic real-team & category-matched car pairings for new championship
                added_count = self.auto_populate_teams_for_championship(selected_cats_list, target_grid_size=40)

            self.refresh_editor_views()
            dlg.destroy()
            msg = f"Championship '{name}' successfully saved to staging library!"
            if mode == "add" and added_count > 0:
                msg += f"\n\nAuto-populated {added_count} real team & car entries across {', '.join(selected_cats_list)}!"
            messagebox.showinfo("Success", msg)

        tk.Button(dlg, text="Save Championship & Calendar", command=save_championship_form, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=20, pady=10)

    def auto_populate_teams_for_championship(self, categories: List[str], target_grid_size: int = 40) -> int:
        """Auto-populates `target_grid_size` (default 40) team-and-car pairings for a newly created championship.
        Strictly harvests real team names from the pre-existing workbook/staged cars pool (never generating
        generic 'Team 1 / Team 2' placeholders) and pairs each real team with a car matching the championship's
        assigned category/categories strictly (using exact matching).
        """
        if not categories:
            return 0

        # 1. Harvest master pool of real team names from staged_cars and engine.cars
        real_teams = []
        seen_teams = set()
        for source_list in (self.staged_cars, engine.cars):
            for entry in source_list:
                t_name = entry.get("team_name", "").strip()
                if t_name and t_name.lower() not in ("factory squad", "team 1", "team 2") and t_name.lower() not in seen_teams:
                    seen_teams.add(t_name.lower())
                    real_teams.append(t_name)

        if not real_teams:
            return 0

        # 2. Build category -> valid car models map from staged_cars and engine.cars using STRICT EXACT MATCHING
        cat_models_map = {}
        for cat in categories:
            cat_clean = cat.strip()
            cat_lower = cat_clean.lower()
            models = []
            for source_list in (self.staged_cars, engine.cars):
                for entry in source_list:
                    e_cat = entry.get("category", "").strip().lower()
                    e_model = entry.get("car_model", "").strip()
                    # FIXED: Strict exact match to prevent 'P1' from matching 'LMP1'
                    if e_model and e_cat == cat_lower:
                        if e_model not in models:
                            models.append(e_model)
            # If a brand-new custom category has no pre-existing cars in the spreadsheet yet
            if not models:
                models = [f"{cat_clean} Spec Car"]
            cat_models_map[cat_clean] = models

        # 3. Track existing (team, category, model) pairings so we prioritize unique combinations first
        existing_pairs = set()
        for entry in self.staged_cars:
            existing_pairs.add((
                entry.get("team_name", "").strip().lower(),
                entry.get("category", "").strip().lower(),
                entry.get("car_model", "").strip().lower()
            ))

        # 4. Shuffle real teams for variety and generate 40 authentic entries across the championship categories
        shuffled_teams = list(real_teams)
        random.shuffle(shuffled_teams)

        added = 0
        cat_list = [c.strip() for c in categories if c.strip()]
        team_idx = 0
        max_attempts = target_grid_size * 5
        attempts = 0

        while added < target_grid_size and attempts < max_attempts:
            attempts += 1
            cat = cat_list[added % len(cat_list)]
            models = cat_models_map[cat]
            team_name = shuffled_teams[team_idx % len(shuffled_teams)]
            car_model = models[(team_idx // len(shuffled_teams) + added) % len(models)]
            team_idx += 1

            pair_key = (team_name.lower(), cat.lower(), car_model.lower())
            # Prioritize unique (real team + category + car model) combinations; allow multi-car entries once unique pairs are exhausted
            if pair_key not in existing_pairs or attempts > len(shuffled_teams) * len(cat_list):
                self.staged_cars.append({
                    "team_name": team_name,
                    "category": cat,
                    "car_model": car_model
                })
                existing_pairs.add(pair_key)
                added += 1

        return added

    def delete_championship(self):
        sel = self.champ_tree.selection()
        if not sel: return
        idx = int(sel[0])
        champ_name = self.staged_championships[idx]['name']
        del self.staged_championships[idx]
        # Also remove associated events
        self.staged_events = [e for e in self.staged_events if e['championship'].lower() != champ_name.lower()]
        self.refresh_editor_views()

    def add_event_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Add Event / Track")
        dlg.geometry("400x300")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        champ_names = [c['name'] for c in self.staged_championships]
        if not champ_names:
            champ_names = ["Default Championship"]

        tk.Label(dlg, text="Select Championship:", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(10, 2))
        c_combo = ttk.Combobox(dlg, values=champ_names, state="readonly", width=38)
        c_combo.pack(padx=15)
        c_combo.set(champ_names[0])

        tk.Label(dlg, text="Track Name:", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(8, 2))
        e_track = ttk.Entry(dlg, width=40)
        e_track.pack(padx=15)

        tk.Label(dlg, text="Layout Variant (e.g. Grand Prix):", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(8, 2))
        e_variant = ttk.Entry(dlg, width=40)
        e_variant.pack(padx=15)
        e_variant.insert(0, "Grand Prix")

        def save_event():
            track = e_track.get().strip()
            if not track:
                messagebox.showwarning("Error", "Track name cannot be empty.")
                return
            self.staged_events.append({
                "championship": c_combo.get(),
                "track_name": track,
                "variant": e_variant.get().strip() or "Grand Prix"
            })
            self.refresh_editor_views()
            dlg.destroy()

        tk.Button(dlg, text="Add Event", command=save_event, bg="#059669", fg="white", font=("Helvetica", 9, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=15, pady=20)

    def delete_event(self):
        sel = self.event_tree.selection()
        if not sel: return
        idx = int(sel[0])
        del self.staged_events[idx]
        self.refresh_editor_views()

    def add_car_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Add Car Model")
        dlg.geometry("400x300")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Team Name:", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(10, 2))
        e_team = ttk.Entry(dlg, width=40)
        e_team.pack(padx=15)

        tk.Label(dlg, text="Car Class Category (Select or type new):", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(8, 2))
        e_cat = ttk.Combobox(dlg, values=self.get_all_available_categories(), width=38)
        e_cat.pack(padx=15)

        tk.Label(dlg, text="Car Model / Vehicle Name:", font=("Helvetica", 8, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=15, pady=(8, 2))
        e_model = ttk.Entry(dlg, width=40)
        e_model.pack(padx=15)

        def save_car():
            model = e_model.get().strip()
            if not model:
                messagebox.showwarning("Error", "Car model cannot be empty.")
                return
            self.staged_cars.append({
                "team_name": e_team.get().strip() or "Factory Squad",
                "category": e_cat.get().strip() or "General",
                "car_model": model
            })
            self.refresh_editor_views()
            dlg.destroy()

        tk.Button(dlg, text="Add Car", command=save_car, bg="#059669", fg="white", font=("Helvetica", 9, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=15, pady=20)

    def delete_car(self):
        sel = self.car_tree.selection()
        if not sel: return
        idx = int(sel[0])
        del self.staged_cars[idx]
        self.refresh_editor_views()

    def save_as_new_blueprint(self):
        f_path = filedialog.asksaveasfilename(initialdir=BLUEPRINTS_DIR, defaultextension=".xlsx", filetypes=[("Excel Blueprints", "*.xlsx")], title="Save New Blueprint As")
        if not f_path:
            return

        try:
            wb = openpyxl.Workbook()
            
            # 1. Championships Sheet
            ws_champ = wb.active
            ws_champ.title = "championships"
            ws_champ.append(["Name", "Categories", "Prestige", "Start_Type", "Field_Type", "Parc_Ferme", "Tags", "Events_Count"])
            for c in self.staged_championships:
                cats_str = ", ".join(c.get('categories', []))
                tags_str = ", ".join(c.get('tags', []))
                ws_champ.append([c['name'], cats_str, c['prestige'], c['start_type'], c['field_type'], c['parc_ferme'], tags_str, c['events_count']])

            # 2. Events Sheet
            ws_events = wb.create_sheet(title="events")
            ws_events.append(["Championship", "Track_Name", "Variant"])
            for e in self.staged_events:
                ws_events.append([e['championship'], e['track_name'], e['variant']])

            # 3. Cars Sheet
            ws_cars = wb.create_sheet(title="cars")
            ws_cars.append(["Team", "Category", "Model"])
            for car in self.staged_cars:
                ws_cars.append([car['team_name'], car['category'], car['car_model']])

            # 4. Categories Sheet (preserves parsed + newly typed custom categories)
            all_cats = self.get_all_available_categories()
            if all_cats:
                ws_cats = wb.create_sheet(title="categories")
                ws_cats.append(["Category"])
                for cat_val in all_cats:
                    ws_cats.append([cat_val])

            wb.save(f_path)
            messagebox.showinfo("Success", f"New blueprint successfully created and saved to:\n{f_path}")
            self.destroy()
            
            # Refresh parent application library
            if hasattr(self.master, "refresh_blueprint_library"):
                self.master.refresh_blueprint_library()
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save blueprint file:\n{e}")


# --- MAIN APPLICATION CLASS ---

class AMS2CareerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AMS2 Career Manager")
        self.geometry("1280x820")
        self.configure(bg="#090d16")
        
        self.parser = SecondMonitorAdvancedParser()
        self.override_mgr = OverrideManager()
        
        self.set_window_icon()
        
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        self.style.configure("Treeview", 
                             background="#1e293b", 
                             foreground="#f8fafc", 
                             fieldbackground="#1e293b",
                             rowheight=26,
                             bordercolor="#334155",
                             lightcolor="#334155",
                             darkcolor="#334155",
                             borderwidth=1)
        self.style.map("Treeview", 
                       background=[("selected", "#2563eb")], 
                       foreground=[("selected", "#ffffff")])
        
        self.style.configure("Treeview.Heading", 
                             background="#0f172a", 
                             foreground="#94a3b8", 
                             font=("Helvetica", 9, "bold"),
                             relief=tk.SOLID,
                             bordercolor="#334155",
                             darkcolor="#334155",
                             lightcolor="#334155",
                             borderwidth=1)
        self.style.map("Treeview.Heading", background=[("active", "#1e293b")])
        
        self.style.configure("TCombobox", fieldbackground="#1e293b", background="#334155", foreground="#ffffff")
        self.style.configure("TSpinbox", fieldbackground="#1e293b", background="#334155", foreground="#ffffff")
        self.style.configure("TEntry", fieldbackground="#1e293b", foreground="#ffffff")

        self.active_profile = None
        self.create_widgets()
        self.load_profiles_list()

    def set_window_icon(self):
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "app_icon.ico")
            else:
                icon_path = os.path.abspath("app_icon.ico")
            
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
                self.update_idletasks()
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                if not hwnd:
                    hwnd = self.winfo_id()
                
                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x00000010
                hicon_large = ctypes.windll.user32.LoadImageW(0, icon_path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
                hicon_small = ctypes.windll.user32.LoadImageW(0, icon_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                
                WM_SETICON = 0x0080
                ICON_SMALL = 0
                ICON_BIG = 1
                
                if hicon_large:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_large)
                if hicon_small:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
        except Exception:
            pass

    def create_widgets(self):
        header = tk.Frame(self, bg="#030712", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(header, text="AMS2 Career Manager", font=("Helvetica", 15, "bold"), fg="#f8fafc", bg="#030712").pack(side=tk.LEFT, padx=20, pady=12)
        
        btn_frame = tk.Frame(header, bg="#030712")
        btn_frame.pack(side=tk.RIGHT, padx=20)
        
        # Version number badge box (styled identically to match button height)
        tk.Button(btn_frame, text="v2.3", state="disabled", disabledforeground="#94a3b8", bg="#1e293b", relief=tk.FLAT, font=("Helvetica", 9, "bold"), padx=10, pady=5).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Career History", command=self.show_history_screen, bg="#334155", fg="white", relief=tk.FLAT, font=("Helvetica", 9, "bold"), padx=10, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Switch Career", command=self.show_select_screen, bg="#1e293b", fg="white", relief=tk.FLAT, font=("Helvetica", 9, "bold"), padx=10, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Import Blueprint (.xlsx)", command=self.import_blueprint_file, bg="#059669", fg="white", relief=tk.FLAT, font=("Helvetica", 9, "bold"), padx=10, pady=5).pack(side=tk.LEFT, padx=5)

        self.container = tk.Frame(self, bg="#090d16")
        self.container.pack(fill=tk.BOTH, expand=True)

        self.screens = {}
        self.build_select_screen()
        self.build_dashboard_screen()
        self.build_summary_screen()
        self.build_history_screen()

        self.show_select_screen()

    def show_screen(self, name):
        for s in self.screens.values(): s.pack_forget()
        self.screens[name].pack(fill=tk.BOTH, expand=True)

    # ------------------ SCREEN 1: CAREER SELECT ------------------
    def build_select_screen(self):
        frame = tk.Frame(self.container, bg="#090d16", padx=30, pady=25)
        self.screens['select'] = frame

        tk.Label(frame, text="Saved Careers", font=("Helvetica", 18, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w")
        
        bp_bar = tk.Frame(frame, bg="#1e293b", padx=10, pady=8)
        bp_bar.pack(fill=tk.X, pady=(6, 12))

        tk.Label(bp_bar, text="Blueprint Library:", font=("Helvetica", 9, "bold"), bg="#1e293b", fg="#94a3b8").pack(side=tk.LEFT, padx=(0, 8))
        
        self.bp_combo_var = tk.StringVar()
        self.bp_combo = ttk.Combobox(bp_bar, textvariable=self.bp_combo_var, state="readonly", width=35)
        self.bp_combo.pack(side=tk.LEFT, padx=5)
        self.bp_combo.bind("<<ComboboxSelected>>", self.on_select_blueprint_from_library)

        tk.Button(bp_bar, text="Refresh Library", command=self.refresh_blueprint_library, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        # Redundant active file text label removed entirely to prevent button squeezing
        
        # In-App Blueprint Editor & Manage Library Buttons (Aligned to the Right)
        tk.Button(bp_bar, text="Blueprint Editor", command=self.open_blueprint_editor, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)
        tk.Button(bp_bar, text="Manage Library", command=self.open_blueprints_folder, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)
        
        self.status_lbl = tk.Label(frame, text="Blueprint Status: Not Loaded", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#ef4444")
        self.status_lbl.pack(anchor="w", pady=(2, 10))

        btn_box = tk.Frame(frame, bg="#090d16")
        btn_box.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(btn_box, text="+ New Career Profile", command=self.open_new_career_dialog, bg="#2563eb", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, relief=tk.FLAT).pack(side=tk.LEFT)

        list_frame = tk.Frame(frame, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(list_frame, columns=("driver", "season", "team", "class", "blueprint"), show="headings", height=10)
        self.tree.heading("driver", text="Driver Name")
        self.tree.heading("season", text="Current Season")
        self.tree.heading("team", text="Current Team / Car")
        self.tree.heading("class", text="Championship Series")
        self.tree.heading("blueprint", text="Source Blueprint")
        
        self.tree.column("driver", width=140, stretch=False)
        self.tree.column("season", width=100, stretch=False)
        self.tree.column("team", width=240, stretch=False)
        self.tree.column("class", width=220, stretch=False)
        self.tree.column("blueprint", width=200, stretch=True)

        self.tree.pack(fill=tk.BOTH, expand=True)

        action_frame = tk.Frame(frame, bg="#090d16", pady=15)
        action_frame.pack(fill=tk.X)
        tk.Button(action_frame, text="Load Selected Career", command=self.load_selected_career, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), padx=20, pady=8, relief=tk.FLAT).pack(side=tk.LEFT)
        tk.Button(action_frame, text="Export Save (JSON)", command=self.export_selected_career, bg="#334155", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=10)
        tk.Button(action_frame, text="Import Save (JSON)", command=self.import_career_save, bg="#334155", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, relief=tk.FLAT).pack(side=tk.LEFT)
        tk.Button(action_frame, text="Delete Selected Career", command=self.delete_selected_career, bg="#dc2626", fg="white", font=("Helvetica", 10, "bold"), padx=20, pady=8, relief=tk.FLAT).pack(side=tk.RIGHT)

        self.refresh_blueprint_library()

    def open_blueprint_editor(self):
        BlueprintEditorWindow(self)

    def export_selected_career(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a career profile to export.")
            return

        career_id = sel[0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, driver_name, season, current_contract, races, calendar, history, blueprint_name, second_monitor_id FROM careers WHERE id = ?", (career_id,))
        r = c.fetchone()
        conn.close()

        if not r:
            return

        profile_data = {
            "id": r[0],
            "name": r[1],
            "driver_name": r[2],
            "season": r[3],
            "current_contract": json.loads(r[4]),
            "races": json.loads(r[5]),
            "calendar": json.loads(r[6]),
            "history": json.loads(r[7]) if r[7] else [],
            "blueprint_name": r[8],
            "second_monitor_id": r[9]
        }

        default_filename = f"{r[1].replace(' ', '_')}_Season_{r[3]}_Backup.json"
        f_path = filedialog.asksaveasfilename(defaultextension=".json", initialfile=default_filename, filetypes=[("JSON Backup Files", "*.json")])
        if f_path:
            try:
                with open(f_path, "w", encoding="utf-8") as f:
                    json.dump(profile_data, f, indent=4)
                messagebox.showinfo("Export Successful", f"Career profile successfully exported to:\n{f_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Could not export profile:\n{e}")

    def import_career_save(self):
        f_path = filedialog.askopenfilename(filetypes=[("JSON Backup Files", "*.json")])
        if f_path:
            try:
                with open(f_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                required_keys = ["id", "name", "driver_name", "season", "current_contract", "races", "calendar"]
                if not all(k in data for k in required_keys):
                    messagebox.showerror("Invalid Format", "The selected JSON file is not a valid AMS2 Career Manager backup.")
                    return

                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                
                c.execute("SELECT id FROM careers WHERE id = ?", (data["id"],))
                exists = c.fetchone()
                if exists:
                    data["id"] = f"prof_{random.randint(10000, 99999)}"

                c.execute("INSERT OR REPLACE INTO careers VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                          (
                              data["id"],
                              data["name"],
                              data["driver_name"],
                              data["season"],
                              json.dumps(data["current_contract"]),
                              json.dumps(data["races"]),
                              json.dumps(data["calendar"]),
                              json.dumps(data.get("history", [])),
                              data.get("blueprint_name", ""),
                              data.get("second_monitor_id", "")
                          ))
                conn.commit()
                conn.close()

                self.load_profiles_list()
                messagebox.showinfo("Import Successful", f"Career profile '{data['name']}' imported successfully!")
            except Exception as e:
                messagebox.showerror("Import Error", f"Could not import career save:\n{e}")

    def open_blueprints_folder(self):
        if not os.path.exists(BLUEPRINTS_DIR):
            os.makedirs(BLUEPRINTS_DIR, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(BLUEPRINTS_DIR)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", BLUEPRINTS_DIR])
            else:
                subprocess.Popen(["xdg-open", BLUEPRINTS_DIR])
        except Exception as e:
            messagebox.showerror("Error Opening Directory", f"Could not open blueprints folder:\n{e}")

    def delete_selected_career(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a career profile to delete.")
            return

        career_id = sel[0]
        item_vals = self.tree.item(career_id, 'values')
        driver_name = item_vals[0] if item_vals else "this profile"

        confirm = messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to permanently delete the career profile for '{driver_name}'?\n\nThis action cannot be undone.",
            icon='warning'
        )

        if confirm:
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("DELETE FROM careers WHERE id = ?", (career_id,))
                conn.commit()
                conn.close()

                if self.active_profile and self.active_profile.get('id') == career_id:
                    self.active_profile = None

                self.load_profiles_list()
                messagebox.showinfo("Deleted", f"Career profile for '{driver_name}' has been deleted.")
            except Exception as e:
                messagebox.showerror("Error Deleting Profile", f"Could not delete career profile:\n{e}")

    def import_blueprint_file(self):
        f_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if f_path:
            try:
                import shutil
                filename = os.path.basename(f_path)
                dest_path = os.path.join(BLUEPRINTS_DIR, filename)
                shutil.copy(f_path, dest_path)
                
                count = engine.load_from_excel(dest_path)
                self.refresh_blueprint_library()
                self.bp_combo.set(filename)
                self.update_blueprint_width_and_label(filename)
                self.status_lbl.config(text=f"Blueprint Status: Loaded '{filename}' ({count} Championships)", fg="#10b981")
                messagebox.showinfo("Success", f"Imported and loaded '{filename}' successfully!")
            except Exception as e:
                messagebox.showerror("Error Importing Excel", f"Could not import blueprint workbook:\n{e}")

    def refresh_blueprint_library(self):
        if not os.path.exists(BLUEPRINTS_DIR):
            os.makedirs(BLUEPRINTS_DIR, exist_ok=True)
        files = [f for f in os.listdir(BLUEPRINTS_DIR) if f.endswith('.xlsx')]
        self.bp_combo['values'] = files
        if files:
            if engine.active_blueprint_filename in files:
                target_file = engine.active_blueprint_filename
            else:
                target_file = files[0]
            self.bp_combo.set(target_file)
            self.load_blueprint_by_filename(target_file)
        else:
            self.bp_combo.set('')
            self.status_lbl.config(text="Blueprint Status: No Blueprints Found in Library (Import above)", fg="#ef4444")

    def update_blueprint_width_and_label(self, filename):
        dynamic_width = max(35, min(len(filename) + 5, 55))
        self.bp_combo.config(width=dynamic_width)

    def on_select_blueprint_from_library(self, event):
        filename = self.bp_combo.get()
        if filename:
            self.load_blueprint_by_filename(filename)

    def load_blueprint_by_filename(self, filename):
        f_path = os.path.join(BLUEPRINTS_DIR, filename)
        if os.path.exists(f_path):
            try:
                count = engine.load_from_excel(f_path)
                self.update_blueprint_width_and_label(filename)
                self.status_lbl.config(text=f"Blueprint Status: Active '{filename}' ({count} Championships)", fg="#10b981")
            except Exception as e:
                messagebox.showerror("Error", f"Could not load blueprint {filename}:\n{e}")

    def load_profiles_list(self):
        for i in self.tree.get_children(): 
            self.tree.delete(i)
            
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, driver_name, season, current_contract, blueprint_name FROM careers")
        rows = c.fetchall()
        conn.close()

        for idx, r in enumerate(rows):
            profile_id = r[0]
            contract = json.loads(r[4])
            bp_used = r[5] if len(r) > 5 and r[5] else "Default"
            row_vals = (r[2], f"Season {r[3]}", f"{contract.get('teamName', 'N/A')} ({contract.get('carModel', 'N/A')})", contract.get('champName', 'N/A'), bp_used)
            
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", tk.END, iid=profile_id, values=row_vals, tags=(tag,))

        self.tree.tag_configure("evenrow", background="#1e293b")
        self.tree.tag_configure("oddrow", background="#162032")

    def show_select_screen(self):
        self.load_profiles_list()
        self.show_screen('select')

    # ------------------ SCREEN 2: UNIFIED DASHBOARD WITH TABBED INSPECTOR ------------------
    def build_dashboard_screen(self):
        frame = tk.Frame(self.container, bg="#090d16", padx=15, pady=10)
        self.screens['dashboard'] = frame

        self.dash_header = tk.Label(frame, text="", font=("Helvetica", 13, "bold"), bg="#090d16", fg="#f8fafc", anchor="w")
        self.dash_header.pack(fill=tk.X, pady=(0, 2))

        # Enhanced Metadata Strip featuring active Car Classes, Blueprint Switcher, and Trajectory Calculator Toggle
        meta_strip_bar = tk.Frame(frame, bg="#1e293b")
        meta_strip_bar.pack(fill=tk.X, pady=(0, 6))

        self.meta_lbl = tk.Label(meta_strip_bar, text="", font=("Helvetica", 9, "bold"), bg="#1e293b", fg="#94a3b8", anchor="w", padx=10, pady=4)
        self.meta_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Subtle Ti System Selector Toggle (Persistent per session)
        ti_toggle_frame = tk.Frame(meta_strip_bar, bg="#1e293b")
        ti_toggle_frame.pack(side=tk.RIGHT, padx=10)

        tk.Label(ti_toggle_frame, text="Ti System:", font=("Helvetica", 8, "bold"), bg="#1e293b", fg="#94a3b8").pack(side=tk.LEFT, padx=(0, 4))
        self.ti_mode_var = tk.StringVar(value="Advanced Ti")
        self.ti_mode_combo = ttk.Combobox(ti_toggle_frame, textvariable=self.ti_mode_var, values=["Legacy Ti (Fixed)", "Advanced Ti (Non-Linear + Recovery)"], state="readonly", width=28)
        self.ti_mode_combo.pack(side=tk.LEFT, padx=4)
        self.ti_mode_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_dashboard())

        tk.Button(meta_strip_bar, text="Update Save Blueprint", command=self.change_save_blueprint, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=8, pady=2, relief=tk.FLAT).pack(side=tk.RIGHT, padx=10)

        # Main Split Frame: Left = Rank Status, Event Calendar & Detailed Championship Classes | Right = Second Monitor XML Integration Tabs
        main_split = tk.Frame(frame, bg="#090d16")
        main_split.pack(fill=tk.BOTH, expand=True)

        # ------------------ LEFT COLUMN: STATUS, CALENDAR & CAR CLASSES ------------------
        left_col = tk.Frame(main_split, bg="#1e293b", width=340, highlightbackground="#334155", highlightthickness=1, padx=10, pady=10)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_col.pack_propagate(False)

        tk.Label(left_col, text="Season Rank Status", font=("Helvetica", 10, "bold"), bg="#1e293b", fg="#f8fafc").pack(anchor="w")
        
        self.rank_box = tk.Frame(left_col, bg="#059669", padx=10, pady=6)
        self.rank_box.pack(fill=tk.X, pady=(4, 6))
        
        self.rank_title_lbl = tk.Label(self.rank_box, text="PROMOTION", font=("Helvetica", 11, "bold"), bg="#059669", fg="white")
        self.rank_title_lbl.pack()
        self.rank_sub_lbl = tk.Label(self.rank_box, text="Trajectory: 85.0%", font=("Helvetica", 9), bg="#059669", fg="white")
        self.rank_sub_lbl.pack(pady=1)

        # Season Event Calendar Panel
        cal_header_row = tk.Frame(left_col, bg="#1e293b")
        cal_header_row.pack(fill=tk.X, pady=(0, 2))
        tk.Label(cal_header_row, text="Season Event Calendar", font=("Helvetica", 9, "bold"), bg="#1e293b", fg="#f8fafc").pack(side=tk.LEFT)
        tk.Button(cal_header_row, text="Edit Track", command=self.edit_selected_calendar_track, bg="#d97706", fg="white", font=("Helvetica", 8, "bold"), padx=6, relief=tk.FLAT).pack(side=tk.RIGHT)
        
        cal_frame = tk.Frame(left_col, bg="#1e293b")
        cal_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 6))
        
        self.cal_listbox = tk.Listbox(cal_frame, font=("Helvetica", 8), bg="#090d16", fg="#f8fafc", selectbackground="#2563eb", selectforeground="#ffffff", borderwidth=0, highlightthickness=0, height=7)
        self.cal_listbox.pack(fill=tk.BOTH, expand=True)

        # Detailed Championship Classes Breakdown Panel
        class_panel_header = tk.Frame(left_col, bg="#1e293b")
        class_panel_header.pack(fill=tk.X, pady=(4, 2))
        tk.Label(class_panel_header, text="Championship Classes & Eligible Cars", font=("Helvetica", 9, "bold"), bg="#1e293b", fg="#38bdf8").pack(side=tk.LEFT)

        class_tree_frame = tk.Frame(left_col, bg="#090d16", highlightbackground="#334155", highlightthickness=1)
        class_tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        cols = ("class_name", "models")
        self.class_details_tree = ttk.Treeview(class_tree_frame, columns=cols, show="headings", height=5)
        self.class_details_tree.heading("class_name", text="Class / Category")
        self.class_details_tree.heading("models", text="Permitted Vehicle Models")

        self.class_details_tree.column("class_name", width=110, stretch=False)
        self.class_details_tree.column("models", width=240, stretch=True)

        class_h_scroll = ttk.Scrollbar(class_tree_frame, orient=tk.HORIZONTAL, command=self.class_details_tree.xview)
        class_v_scroll = ttk.Scrollbar(class_tree_frame, orient=tk.VERTICAL, command=self.class_details_tree.yview)
        self.class_details_tree.configure(xscrollcommand=class_h_scroll.set, yscrollcommand=class_v_scroll.set)

        self.class_details_tree.grid(row=0, column=0, sticky="nsew")
        class_v_scroll.grid(row=0, column=1, sticky="ns")
        class_h_scroll.grid(row=1, column=0, sticky="ew")

        class_tree_frame.grid_rowconfigure(0, weight=1)
        class_tree_frame.grid_columnconfigure(0, weight=1)

        tk.Button(left_col, text="End Season & Evaluate Offers", command=self.evaluate_season, bg="#0f172a", fg="white", font=("Helvetica", 9, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, side=tk.BOTTOM)

        # ------------------ RIGHT COLUMN: TABBED SECOND MONITOR XML & ROUND RECORDER ------------------
        right_col = tk.Frame(main_split, bg="#090d16")
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # XML Control Bar
        xml_ctrl = tk.Frame(right_col, bg="#1e293b", padx=10, pady=6)
        xml_ctrl.pack(fill=tk.X, pady=(0, 6))

        tk.Button(xml_ctrl, text="Link Series", command=self.open_link_series_dialog, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(xml_ctrl, text="Reload XML", command=self.refresh_xml_data, bg="#059669", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(xml_ctrl, text="Set XML Path", command=self.configure_xml_path, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        
        self.linked_series_lbl = tk.Label(xml_ctrl, text="Linked ID: None", font=("Helvetica", 8, "bold"), bg="#1e293b", fg="#10b981")
        self.linked_series_lbl.pack(side=tk.LEFT, padx=10)

        # TABBED NOTEBOOK FOR XML VIEWS
        self.xml_notebook = ttk.Notebook(right_col)
        self.xml_notebook.pack(fill=tk.BOTH, expand=True)

        # TAB 1: Championship Standings Table (Live XML with Class Scope Filter)
        self.tab_standings = ttk.Frame(self.xml_notebook)
        self.xml_notebook.add(self.tab_standings, text="Championship Standings (Live XML)")

        standings_top_bar = tk.Frame(self.tab_standings, bg="#1e293b", padx=8, pady=4)
        standings_top_bar.pack(fill=tk.X)
        tk.Label(standings_top_bar, text="Standings Scope:", bg="#1e293b", fg="#94a3b8", font=("Helvetica", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.standings_scope_combo = ttk.Combobox(standings_top_bar, values=["Overall Standings"], state="readonly", width=22)
        self.standings_scope_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.standings_scope_combo.set("Overall Standings")
        self.standings_scope_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_current_championship_view())
        
        columns = ("pos", "driver", "car", "class_name", "points")
        self.standings_tree = ttk.Treeview(self.tab_standings, columns=columns, show="headings", height=8)
        self.standings_tree.heading("pos", text="Pos")
        self.standings_tree.heading("driver", text="Driver Name")
        self.standings_tree.heading("car", text="Car / Model")
        self.standings_tree.heading("class_name", text="Class")
        self.standings_tree.heading("points", text="Points")

        self.standings_tree.column("pos", width=40, anchor="center")
        self.standings_tree.column("driver", width=180, anchor="w")
        self.standings_tree.column("car", width=160, anchor="w")
        self.standings_tree.column("class_name", width=90, anchor="center")
        self.standings_tree.column("points", width=70, anchor="center")

        s_scroll = ttk.Scrollbar(self.tab_standings, orient=tk.VERTICAL, command=self.standings_tree.yview)
        self.standings_tree.configure(yscrollcommand=s_scroll.set)
        
        self.standings_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        s_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 2: Per-Round Classification Inspector & Hybrid Result Recorder
        self.tab_round_inspector = ttk.Frame(self.xml_notebook)
        self.xml_notebook.add(self.tab_round_inspector, text="Per-Round Classification & Result Recorder")

        r_ctrl = tk.Frame(self.tab_round_inspector, bg="#1e293b", padx=8, pady=6)
        r_ctrl.pack(fill=tk.X)

        tk.Label(r_ctrl, text="Select Round:", bg="#1e293b", fg="#94a3b8", font=("Helvetica", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.xml_round_combo = ttk.Combobox(r_ctrl, state="readonly", width=18)
        self.xml_round_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.xml_round_combo.bind("<<ComboboxSelected>>", self.on_xml_round_select)

        tk.Label(r_ctrl, text="Round Scope:", bg="#1e293b", fg="#94a3b8", font=("Helvetica", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.round_scope_combo = ttk.Combobox(r_ctrl, values=["All Classes"], state="readonly", width=18)
        self.round_scope_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.round_scope_combo.set("All Classes")
        self.round_scope_combo.bind("<<ComboboxSelected>>", lambda e: self.on_xml_round_select(None))

        tk.Button(r_ctrl, text="Edit Round XML", command=self.open_round_editor, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=6, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        self.multiclass_lbl = tk.Label(r_ctrl, text="Class Mode: Standard", font=("Helvetica", 8, "bold"), bg="#1e293b", fg="#38bdf8")
        self.multiclass_lbl.pack(side=tk.RIGHT, padx=5)

        round_grid_frame = tk.Frame(self.tab_round_inspector, bg="#090d16")
        round_grid_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        r_cols = ("pos", "class_pos", "driver", "car", "class_name", "points", "fastest_lap")
        self.xml_round_tree = ttk.Treeview(round_grid_frame, columns=r_cols, show="headings", height=6)
        self.xml_round_tree.heading("pos", text="Ovr Pos")
        self.xml_round_tree.heading("class_pos", text="Cls Pos")
        self.xml_round_tree.heading("driver", text="Driver Name")
        self.xml_round_tree.heading("car", text="Car Model")
        self.xml_round_tree.heading("class_name", text="Class")
        self.xml_round_tree.heading("points", text="Points")
        self.xml_round_tree.heading("fastest_lap", text="FL")

        self.xml_round_tree.column("pos", width=55, anchor="center")
        self.xml_round_tree.column("class_pos", width=55, anchor="center")
        self.xml_round_tree.column("driver", width=150, anchor="w")
        self.xml_round_tree.column("car", width=130, anchor="w")
        self.xml_round_tree.column("class_name", width=80, anchor="center")
        self.xml_round_tree.column("points", width=60, anchor="center")
        self.xml_round_tree.column("fastest_lap", width=40, anchor="center")

        r_scroll = ttk.Scrollbar(round_grid_frame, orient=tk.VERTICAL, command=self.xml_round_tree.yview)
        self.xml_round_tree.configure(yscrollcommand=r_scroll.set)

        self.xml_round_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        r_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        input_grid = tk.Frame(self.tab_round_inspector, bg="#1e293b", padx=8, pady=6)
        input_grid.pack(fill=tk.X)

        tk.Label(input_grid, text="Quali Pos:", bg="#1e293b", fg="#94a3b8", font=("Helvetica", 8, "bold")).grid(row=0, column=0, padx=4, sticky="w")
        self.lock_q = ttk.Spinbox(input_grid, from_=1, to=100, width=6)
        self.lock_q.grid(row=0, column=1, padx=4)
        self.lock_q.set(1)

        self.qualy_status_lbl = tk.Label(input_grid, text="", font=("Helvetica", 7, "italic"), bg="#1e293b", fg="#38bdf8")
        self.qualy_status_lbl.grid(row=0, column=2, padx=4, sticky="w")

        tk.Label(input_grid, text="Grid Size:", bg="#1e293b", fg="#94a3b8", font=("Helvetica", 8, "bold")).grid(row=0, column=3, padx=4, sticky="w")
        self.lock_grid = ttk.Spinbox(input_grid, from_=2, to=60, width=6)
        self.lock_grid.grid(row=0, column=4, padx=4)
        self.lock_grid.set(20)

        tk.Button(input_grid, text="Lock In & Record Round", command=self.lock_in_xml_round_result, bg="#2563eb", fg="white", font=("Helvetica", 8, "bold"), padx=10, relief=tk.FLAT).grid(row=0, column=5, padx=12)

        logged_frame = tk.Frame(right_col, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        logged_frame.pack(fill=tk.X, pady=(6, 0))

        tk.Label(logged_frame, text=" Season Trajectory Logged Results ", font=("Helvetica", 8, "bold"), bg="#1e293b", fg="#94a3b8").pack(anchor="w", padx=6, pady=2)

        self.race_tree = ttk.Treeview(logged_frame, columns=("event", "q", "f", "c", "grid"), show="headings", height=4)
        self.race_tree.heading("event", text="Track & Layout Variant")
        self.race_tree.heading("q", text="Quali")
        self.race_tree.heading("f", text="Finish (Class)")
        self.race_tree.heading("c", text="Champ Pos")
        self.race_tree.heading("grid", text="Grid")
        
        self.race_tree.column("event", width=180, stretch=True)
        self.race_tree.column("q", width=45, anchor="center")
        self.race_tree.column("f", width=75, anchor="center")
        self.race_tree.column("c", width=65, anchor="center")
        self.race_tree.column("grid", width=45, anchor="center")

        self.race_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        logged_actions = tk.Frame(logged_frame, bg="#1e293b", pady=4)
        logged_actions.pack(fill=tk.X)
        tk.Button(logged_actions, text="Delete Selected Result", command=self.delete_selected_race, bg="#dc2626", fg="white", font=("Helvetica", 8, "bold"), padx=6, relief=tk.FLAT).pack(side=tk.LEFT, padx=6)

    def change_save_blueprint(self):
        """Allows the user to re-bind or switch the active source blueprint file for the currently loaded career save."""
        if not self.active_profile:
            messagebox.showwarning("No Active Profile", "Please load a career profile first.")
            return

        dlg = tk.Toplevel(self)
        dlg.title("Update Save Source Blueprint")
        dlg.geometry("520x260")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Re-assign Source Blueprint for Active Save", font=("Helvetica", 11, "bold"), bg="#090d16", fg="#f8fafc").pack(padx=15, pady=(15, 4), anchor="w")
        tk.Label(dlg, text=f"Currently Bound: {self.active_profile.get('blueprint_name', 'None')}", font=("Helvetica", 8, "italic"), bg="#090d16", fg="#38bdf8").pack(padx=15, pady=(0, 10), anchor="w")

        files = [f for f in os.listdir(BLUEPRINTS_DIR) if f.endswith('.xlsx')] if os.path.exists(BLUEPRINTS_DIR) else []

        input_frame = tk.Frame(dlg, bg="#1e293b", padx=12, pady=12)
        input_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        tk.Label(input_frame, text="Select Blueprint from Library:", font=("Helvetica", 9, "bold"), bg="#1e293b", fg="#94a3b8").pack(anchor="w", pady=(0, 5))

        combo_var = tk.StringVar()
        bp_select_combo = ttk.Combobox(input_frame, textvariable=combo_var, values=files, state="readonly", width=45)
        bp_select_combo.pack(anchor="w", pady=(0, 10))

        if files:
            current_bp = self.active_profile.get('blueprint_name', '')
            if current_bp in files:
                bp_select_combo.set(current_bp)
            else:
                bp_select_combo.set(files[0])

        def browse_and_import():
            f_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if f_path:
                try:
                    import shutil
                    filename = os.path.basename(f_path)
                    dest_path = os.path.join(BLUEPRINTS_DIR, filename)
                    shutil.copy(f_path, dest_path)
                    
                    self.refresh_blueprint_library()
                    files_new = [f for f in os.listdir(BLUEPRINTS_DIR) if f.endswith('.xlsx')]
                    bp_select_combo['values'] = files_new
                    bp_select_combo.set(filename)
                    messagebox.showinfo("Imported", f"Successfully imported '{filename}' into your library!")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not import blueprint file:\n{e}")

        tk.Button(input_frame, text="+ Import New Blueprint File", command=browse_and_import, bg="#334155", fg="white", font=("Helvetica", 8, "bold"), padx=8, relief=tk.FLAT).pack(anchor="w")

        def confirm_save_blueprint_change():
            chosen_file = combo_var.get().strip()
            if not chosen_file:
                messagebox.showwarning("Selection Required", "Please select a valid blueprint file from the dropdown.")
                return

            full_path = os.path.join(BLUEPRINTS_DIR, chosen_file)
            if not os.path.exists(full_path):
                messagebox.showerror("File Missing", f"The file '{chosen_file}' could not be found.")
                return

            try:
                count = engine.load_from_excel(full_path)
                self.active_profile['blueprint_name'] = chosen_file
                self.save_active_profile()

                self.bp_combo.set(chosen_file)
                self.update_blueprint_width_and_label(chosen_file)
                self.status_lbl.config(text=f"Blueprint Status: Active '{chosen_file}' ({count} Championships)", fg="#10b981")

                self.refresh_dashboard()
                dlg.destroy()

                # --- ENHANCED CONFIRMATION DIALOG WITH NEW CONTENT LIST ---
                content_dlg = tk.Toplevel(self)
                content_dlg.title("Blueprint Re-bound & Content Summary")
                content_dlg.geometry("700x500")
                content_dlg.configure(bg="#090d16")
                content_dlg.transient(self)
                content_dlg.grab_set()

                tk.Label(content_dlg, text="Blueprint Successfully Re-Bound!", font=("Helvetica", 12, "bold"), bg="#090d16", fg="#10b981").pack(padx=20, pady=(15, 4), anchor="w")
                tk.Label(content_dlg, text=f"Profile '{self.active_profile['name']}' is now linked to: '{chosen_file}'", font=("Helvetica", 9), bg="#090d16", fg="#f8fafc").pack(padx=20, pady=(0, 10), anchor="w")
                tk.Label(content_dlg, text=f"Loaded Content Summary ({count} Championships & {len(engine.cars)} Cars available):", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#38bdf8").pack(padx=20, pady=(0, 5), anchor="w")

                tree_frame = tk.Frame(content_dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
                tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

                cols = ("champ", "prestige", "categories", "events")
                content_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
                content_tree.heading("champ", text="Championship Series")
                content_tree.heading("prestige", text="Prestige")
                content_tree.heading("categories", text="Car Categories")
                content_tree.heading("events", text="Events")

                content_tree.column("champ", width=260, stretch=False)
                content_tree.column("prestige", width=80, anchor="center", stretch=False)
                content_tree.column("categories", width=200, stretch=True)
                content_tree.column("events", width=60, anchor="center", stretch=False)

                v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=content_tree.yview)
                content_tree.configure(yscrollcommand=v_scroll.set)

                content_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

                for idx, c_item in enumerate(engine.championships):
                    cats_str = ", ".join(c_item['categories']) if c_item['categories'] else "General"
                    tag = "evenrow" if idx % 2 == 0 else "oddrow"
                    content_tree.insert("", tk.END, values=(c_item['name'], c_item['prestige'], cats_str, c_item['events_count']), tags=(tag,))

                content_tree.tag_configure("evenrow", background="#1e293b")
                content_tree.tag_configure("oddrow", background="#162032")

                tk.Label(content_dlg, text="New contract offers at end-of-season will now pull live from this updated blueprint!", font=("Helvetica", 8, "italic"), bg="#090d16", fg="#94a3b8").pack(padx=20, pady=(0, 10), anchor="w")
                tk.Button(content_dlg, text="OK", command=content_dlg.destroy, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=20, pady=(0, 15))

            except Exception as e:
                messagebox.showerror("Error", f"Could not load selected blueprint file:\n{e}")

        tk.Button(dlg, text="Apply & Re-bind Save Profile", command=confirm_save_blueprint_change, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=15, pady=(0, 15))

    def edit_selected_calendar_track(self):
        if not self.active_profile:
            return
        
        sel = self.cal_listbox.curselection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a round from the Season Event Calendar to edit.")
            return

        idx = sel[0]
        current_entry = self.active_profile['calendar'][idx]

        dlg = tk.Toplevel(self)
        dlg.title(f"Edit Track for Round {idx+1}")
        dlg.geometry("440x220")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text=f"Edit Track / Layout for Round {idx+1}:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(padx=15, pady=(15, 2), anchor="w")
        tk.Label(dlg, text="Select from blueprint library or type custom name (session-only persistence):", font=("Helvetica", 8), bg="#090d16", fg="#94a3b8").pack(padx=15, pady=(0, 8), anchor="w")

        blueprint_tracks = sorted(list(set([f"{e['track_name']} ({e['variant']})" for e in engine.events])))
        if not blueprint_tracks:
            blueprint_tracks = ["Interlagos (Grand Prix)", "Monza (Grand Prix)", "Spa-Francorchamps (GP)", "Silverstone (GP)", "Nürburgring (GP)"]

        combo_var = tk.StringVar(value=current_entry)
        track_combo = ttk.Combobox(dlg, textvariable=combo_var, values=blueprint_tracks, width=50)
        track_combo.pack(padx=15, pady=5)

        def save_track_edit():
            new_val = combo_var.get().strip()
            if not new_val:
                messagebox.showwarning("Invalid", "Track name cannot be empty.")
                return
            
            self.active_profile['calendar'][idx] = new_val
            self.save_active_profile()
            self.refresh_dashboard()
            dlg.destroy()
            messagebox.showinfo("Updated", f"Round {idx+1} calendar track updated successfully for this career session.")

        tk.Button(dlg, text="Save Track Override", command=save_track_edit, bg="#059669", fg="white", font=("Helvetica", 9, "bold"), pady=6, relief=tk.FLAT).pack(pady=15)

    def open_link_series_dialog(self):
        if not self.active_profile:
            messagebox.showwarning("No Active Profile", "Please load a career profile first.")
            return

        parsed = self.parser.parse_championships()
        if not parsed:
            messagebox.showwarning("No XML Data", f"Could not find or parse any championships from:\n{self.parser.xml_path}")
            return

        dlg = tk.Toplevel(self)
        dlg.title("Link Second Monitor Championship Series")
        dlg.geometry("640x380")
        dlg.configure(bg="#090d16")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Select Second Monitor Championship to Link", font=("Helvetica", 11, "bold"), bg="#090d16", fg="#f8fafc").pack(padx=15, pady=(15, 5), anchor="w")
        tk.Label(dlg, text="This will lock your active career profile to this exact Second Monitor series ID.", font=("Helvetica", 8), bg="#090d16", fg="#94a3b8").pack(padx=15, pady=(0, 10), anchor="w")

        tree_frame = tk.Frame(dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        cols = ("name", "rounds", "id")
        picker_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        picker_tree.heading("name", text="Championship Name in XML")
        picker_tree.heading("rounds", text="Rounds")
        picker_tree.heading("id", text="Global ID")

        picker_tree.column("name", width=260, stretch=False)
        picker_tree.column("rounds", width=80, anchor="center", stretch=False)
        picker_tree.column("id", width=240, stretch=True)

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=picker_tree.yview)
        picker_tree.configure(yscrollcommand=v_scroll.set)

        picker_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        champ_ids = list(parsed.keys())
        for idx, cid in enumerate(champ_ids):
            cdata = parsed[cid]
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            picker_tree.insert("", tk.END, iid=str(idx), values=(cdata.championship_name, len(cdata.rounds), cid), tags=(tag,))

        picker_tree.tag_configure("evenrow", background="#1e293b")
        picker_tree.tag_configure("oddrow", background="#162032")

        if champ_ids:
            picker_tree.selection_set("0")

        def confirm_link():
            sel = picker_tree.selection()
            if not sel:
                messagebox.showwarning("Selection Required", "Please select a championship from the list.")
                return
            
            chosen_cid = champ_ids[int(sel[0])]
            chosen_name = parsed[chosen_cid].championship_name

            self.active_profile["second_monitor_id"] = chosen_cid
            self.save_active_profile()

            messagebox.showinfo("Linked!", f"Successfully linked career profile to:\n'{chosen_name}'")
            dlg.destroy()
            self.refresh_current_championship_view()

        tk.Button(dlg, text="Confirm & Bind Link", command=confirm_link, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=15, pady=(0, 15))

    def configure_xml_path(self):
        filepath = filedialog.askopenfilename(title="Select Second Monitor Championships.xml", filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
        if filepath:
            self.parser.xml_path = Path(filepath)
            self.refresh_xml_data()

    def refresh_xml_data(self):
        self.parsed_champs = self.parser.parse_championships()
        self.refresh_current_championship_view()

    def refresh_current_championship_view(self):
        if not self.active_profile:
            return

        linked_id = self.active_profile.get('second_monitor_id', '')
        contract = self.active_profile.get('currentContract', {})
        active_champ_name = contract.get('champName', '').lower()

        parsed_data = getattr(self, 'parsed_champs', {})
        if not parsed_data:
            parsed_data = self.parser.parse_championships()
            self.parsed_champs = parsed_data

        matched_champ_id = None
        if linked_id and linked_id in parsed_data:
            matched_champ_id = linked_id
        else:
            for cid, cdata in parsed_data.items():
                if active_champ_name in cdata.championship_name.lower() or cdata.championship_name.lower() in active_champ_name:
                    matched_champ_id = cid
                    break
        
        if not matched_champ_id and parsed_data:
            matched_champ_id = list(parsed_data.keys())[0]

        self.current_parsed_champ_id = matched_champ_id

        if linked_id:
            disp_id = linked_id[:12] + "..." if len(linked_id) > 12 else linked_id
            self.linked_series_lbl.config(text=f"Linked ID: {disp_id}", fg="#10b981")
        else:
            self.linked_series_lbl.config(text="Linked ID: None (Auto)", fg="#d97706")

        for item in self.standings_tree.get_children():
            self.standings_tree.delete(item)

        if not matched_champ_id or matched_champ_id not in parsed_data:
            return

        champ: ChampionshipData = parsed_data[matched_champ_id]
        
        scope_options = ["Overall Standings"]
        for c_name in champ.class_standings.keys():
            scope_options.append(f"Class: {c_name}")
        self.standings_scope_combo['values'] = scope_options
        
        current_scope = self.standings_scope_combo.get()
        if not current_scope or current_scope not in scope_options:
            current_scope = "Overall Standings"
            self.standings_scope_combo.set(current_scope)

        # Update round scope combo options as well
        round_scope_options = ["All Classes"] + list(champ.class_standings.keys())
        self.round_scope_combo['values'] = round_scope_options
        if self.round_scope_combo.get() not in round_scope_options:
            self.round_scope_combo.set("All Classes")

        if current_scope == "Overall Standings":
            standings_list = champ.overall_standings
            for st in standings_list:
                self.standings_tree.insert("", tk.END, values=(st.position, st.driver_name, "", "Overall", f"{st.total_points:.1f}"))
        else:
            target_class = current_scope.replace("Class: ", "").strip()
            class_st_list = champ.class_standings.get(target_class, [])
            for st in class_st_list:
                self.standings_tree.insert("", tk.END, values=(st.position, st.driver_name, "", st.class_name, f"{st.total_points:.1f}"))

        round_labels = [f"R{rd['round_number']}: {rd['track_name']}" for rd in champ.rounds]
        self.xml_round_combo["values"] = round_labels
        if round_labels:
            if self.xml_round_combo.current() < 0:
                self.xml_round_combo.current(0)
            self.on_xml_round_select(None)

    def on_xml_round_select(self, event):
        champ_id = getattr(self, 'current_parsed_champ_id', None)
        if not champ_id or champ_id not in getattr(self, 'parsed_champs', {}):
            return
        champ: ChampionshipData = self.parsed_champs[champ_id]
        sel_idx = self.xml_round_combo.current()
        if sel_idx < 0 or sel_idx >= len(champ.rounds):
            return

        round_data = champ.rounds[sel_idx]
        overrides = self.override_mgr.load_overrides(champ_id)
        r_key = str(sel_idx)

        results = None
        if r_key in overrides:
            results = overrides[r_key]
        else:
            for sess in round_data["sessions"]:
                if "race" in sess["session_type"].lower():
                    results = sess["results"]
                    break
            if not results and round_data["sessions"]:
                results = round_data["sessions"][0]["results"]

        for item in self.xml_round_tree.get_children():
            self.xml_round_tree.delete(item)

        if results:
            unique_classes = set()
            for res in results:
                c_name = res.class_name if hasattr(res, 'class_name') else res.get("class_name", "General")
                unique_classes.add(c_name)

            if len(unique_classes) > 1:
                self.multiclass_lbl.config(text=f"Class Mode: Multiclass ({len(unique_classes)} Classes)", fg="#10b981")
            else:
                self.multiclass_lbl.config(text="Class Mode: Single Class", fg="#38bdf8")

            # Group and recalculate class-specific positions and points for this round using scoring rules matrix
            class_groups = {}
            for res in results:
                cn = res.class_name if hasattr(res, 'class_name') else res.get("class_name", "General")
                if cn not in class_groups: 
                    class_groups[cn] = []
                class_groups[cn].append(res)

            processed_results = []
            for cn, group in class_groups.items():
                group.sort(key=lambda x: (x.finish_position if hasattr(x, 'finish_position') else x.get('finish_position', 99)))
                for c_pos, res in enumerate(group, start=1):
                    # Recalculate independent class points using scoring rules
                    matrix = champ.scoring_rules if champ.scoring_rules else [350, 320, 300, 280, 260, 250, 240, 230, 220, 210]
                    recalc_pts = matrix[c_pos - 1] if (c_pos - 1) < len(matrix) else max(5.0, matrix[-1] - ((c_pos - len(matrix)) * 5))
                    
                    if hasattr(res, 'class_finish_position'):
                        res.class_finish_position = c_pos
                        res.points_gain = recalc_pts
                    else:
                        res['class_finish_position'] = c_pos
                        res['points_gain'] = recalc_pts
                    processed_results.append(res)

            def get_pos(r):
                if hasattr(r, 'finish_position'): return r.finish_position
                return int(r.get("finish_position", 99))
            
            sorted_res = sorted(processed_results, key=get_pos)
            
            selected_round_scope = self.round_scope_combo.get()
            
            for res in sorted_res:
                ovr_pos = res.finish_position if hasattr(res, 'finish_position') else res.get("finish_position", 0)
                cls_pos = res.class_finish_position if hasattr(res, 'class_finish_position') else res.get("class_finish_position", 0)
                d_name = res.driver_name if hasattr(res, 'driver_name') else res.get("driver_name", "Unknown")
                c_model = res.car_name if hasattr(res, 'car_name') else res.get("car_name", "")
                cls_name = res.class_name if hasattr(res, 'class_name') else res.get("class_name", "General")
                pts = res.points_gain if hasattr(res, 'points_gain') else res.get("points_gain", 0.0)
                fl = res.fastest_lap if hasattr(res, 'fastest_lap') else res.get("fastest_lap", False)

                if selected_round_scope != "All Classes" and cls_name != selected_round_scope:
                    continue

                self.xml_round_tree.insert("", tk.END, values=(
                    ovr_pos,
                    cls_pos,
                    d_name,
                    c_model,
                    cls_name,
                    pts,
                    "Yes" if fl else "No"
                ))

            detected_grid_size = len(results)
            if detected_grid_size >= 2:
                self.lock_grid.set(detected_grid_size)

            driver_name = self.active_profile.get('driverName', '').lower()

            qualy_pos_found = None
            for sess in round_data["sessions"]:
                if "quali" in sess["session_type"].lower():
                    for q_res in sess["results"]:
                        q_name = q_res.driver_name if hasattr(q_res, 'driver_name') else q_res.get('driver_name', '')
                        is_ply = q_res.is_player if hasattr(q_res, 'is_player') else q_res.get('is_player', False)
                        if is_ply or (driver_name and driver_name in q_name.lower()):
                            qualy_pos_found = q_res.finish_position if hasattr(q_res, 'finish_position') else int(q_res.get('finish_position', 1))
                            break
                    break

            if qualy_pos_found and qualy_pos_found > 0:
                self.lock_q.set(qualy_pos_found)
                self.qualy_status_lbl.config(text="(Auto-Detected from Quali)", fg="#10b981")
            else:
                self.qualy_status_lbl.config(text="(No Qualy Detected - Manual)", fg="#d97706")

    def lock_in_xml_round_result(self):
        if not self.active_profile:
            return
        
        sel_idx = self.xml_round_combo.current()
        if sel_idx < 0:
            messagebox.showwarning("Round Required", "Please select a round from the Per-Round tab dropdown.")
            return

        champ_id = getattr(self, 'current_parsed_champ_id', None)
        if not champ_id or champ_id not in getattr(self, 'parsed_champs', {}):
            messagebox.showwarning("XML Missing", "No Second Monitor championship loaded or matched to parse results.")
            return

        champ: ChampionshipData = self.parsed_champs[champ_id]
        if sel_idx >= len(champ.rounds):
            return

        round_data = champ.rounds[sel_idx]
        r_num = round_data['round_number']
        round_tag = f"R{r_num}"

        existing_rounds = [r['event'].split(":")[0] for r in self.active_profile['races']]
        if round_tag in existing_rounds:
            messagebox.showwarning("Already Locked", f"Round {r_num} has already been locked in for this season.")
            return

        overrides = self.override_mgr.load_overrides(champ_id)
        r_key = str(sel_idx)
        results = overrides.get(r_key)
        if not results:
            for sess in round_data["sessions"]:
                if "race" in sess["session_type"].lower():
                    results = sess["results"]
                    break
            if not results and round_data["sessions"]:
                results = round_data["sessions"][0]["results"]

        if not results:
            messagebox.showwarning("No Results", f"No race results found in XML for Round {r_num}.")
            return

        driver_name = self.active_profile.get('driverName', '').lower()
        player_result = None

        for res in results:
            d_name = res.driver_name if hasattr(res, 'driver_name') else res.get('driver_name', '')
            is_ply = res.is_player if hasattr(res, 'is_player') else res.get('is_player', False)
            if is_ply or (driver_name and driver_name in d_name.lower()):
                player_result = res
                break

        if not player_result and results:
            player_result = results[0]

        if hasattr(player_result, 'class_finish_position'):
            player_class_finish = player_result.class_finish_position
            player_class = player_result.class_name
            ovr_finish = player_result.finish_position
        else:
            player_class = player_result.get('class_name', 'General')
            player_class_finish = int(player_result.get('class_finish_position') or player_result.get('finish_position', 1))
            ovr_finish = int(player_result.get('finish_position', 1))

        if player_class_finish < 1:
            player_class_finish = ovr_finish

        # Determine if this is multiclass to set Champ Pos correctly
        unique_classes = set()
        for res in results:
            c_name = res.class_name if hasattr(res, 'class_name') else res.get("class_name", "General")
            unique_classes.add(c_name)

        is_multiclass = len(unique_classes) > 1

        # For multiclass series, champ pos should be the class finish position instead of overall scratch finish
        player_champ_pos = player_class_finish if is_multiclass else ovr_finish

        try:
            q_pos = int(self.lock_q.get())
            grid_size = int(self.lock_grid.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numeric values for Quali and Grid Size.")
            return

        # For multiclass series, base the grid size on the class-specific field size rather than overall scratch field size
        if is_multiclass:
            class_grid_size = sum(1 for res in results if (res.class_name if hasattr(res, 'class_name') else res.get("class_name", "General")) == player_class)
            grid_size = class_grid_size if class_grid_size > 0 else grid_size

        cal_tracks = self.active_profile.get('calendar', [])
        track_name = cal_tracks[sel_idx] if sel_idx < len(cal_tracks) else round_data['track_name']
        event_str = f"R{r_num}: {track_name}"

        race_record = {
            "event": event_str,
            "q": q_pos,
            "f": player_class_finish,
            "c": player_champ_pos,
            "grid": grid_size
        }

        self.active_profile['races'].append(race_record)
        self.save_active_profile()
        self.refresh_dashboard()
        messagebox.showinfo("Locked In!", f"Successfully locked in Round {r_num} results!\nClass Finish ({player_class}): P{player_class_finish}\nChamp Pos: P{player_champ_pos}\nGrid Used: {grid_size}\nTrajectory Index updated.")

    def open_round_editor(self):
        champ_id = getattr(self, 'current_parsed_champ_id', None)
        if not champ_id or champ_id not in getattr(self, 'parsed_champs', {}):
            messagebox.showwarning("Notice", "No Second Monitor championship series matched or loaded.")
            return
        
        sel_idx = self.xml_round_combo.current()
        if sel_idx < 0:
            messagebox.showwarning("Notice", "Please select a round to edit.")
            return

        champ: ChampionshipData = self.parsed_champs[champ_id]
        round_data = champ.rounds[sel_idx]
        overrides = self.override_mgr.load_overrides(champ_id)
        r_key = str(sel_idx)

        results = overrides.get(r_key)
        if not results:
            raw_results = None
            for sess in round_data["sessions"]:
                if "race" in sess["session_type"].lower():
                    raw_results = sess["results"]
                    break
            if not raw_results and round_data["sessions"]:
                raw_results = round_data["sessions"][0]["results"]
            
            # Convert raw RaceResultEntry dataclass instances into JSON-serializable dictionaries
            results = []
            if raw_results:
                for r in raw_results:
                    if hasattr(r, '__dataclass_fields__') or hasattr(r, 'driver_name'):
                        results.append({
                            "driver_name": getattr(r, 'driver_name', 'Unknown'),
                            "car_name": getattr(r, 'car_name', ''),
                            "class_id": getattr(r, 'class_id', ''),
                            "class_name": getattr(r, 'class_name', 'General'),
                            "finish_position": getattr(r, 'finish_position', 0),
                            "class_finish_position": getattr(r, 'class_finish_position', 0),
                            "total_points": getattr(r, 'total_points', 0.0),
                            "points_gain": getattr(r, 'points_gain', 0.0),
                            "is_player": getattr(r, 'is_player', False),
                            "fastest_lap": getattr(r, 'fastest_lap', False)
                        })
                    else:
                        results.append(dict(r) if isinstance(r, dict) else r)
        else:
            results = [dict(r) if not isinstance(r, dict) else r for r in results]

        if not results:
            messagebox.showinfo("Info", "No results found for this round.")
            return

        editor_win = tk.Toplevel(self)
        editor_win.title(f"Edit Round {round_data['round_number']} - {round_data['track_name']}")
        editor_win.geometry("780x520")
        editor_win.configure(bg="#090d16")

        tk.Label(editor_win, text="Edit Finish Positions, Points, and Fastest Lap Overrides", font=("Helvetica", 11, "bold"), bg="#090d16", fg="#f8fafc").pack(padx=15, pady=10, anchor="w")

        cols = ("driver", "position", "class_position", "points", "fastest_lap")
        tree = ttk.Treeview(editor_win, columns=cols, show="headings", height=12)
        tree.heading("driver", text="Driver Name")
        tree.heading("position", text="Ovr Pos")
        tree.heading("class_position", text="Cls Pos")
        tree.heading("points", text="Points Gain")
        tree.heading("fastest_lap", text="Fastest Lap")

        tree.column("driver", width=220)
        tree.column("position", width=80, anchor="center")
        tree.column("class_position", width=80, anchor="center")
        tree.column("points", width=90, anchor="center")
        tree.column("fastest_lap", width=120, anchor="center")

        tree.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        for i, res in enumerate(results):
            d_name = res.get("driver_name")
            f_pos = res.get("finish_position")
            c_pos = res.get("class_finish_position")
            pts = res.get("points_gain")
            fl = res.get("fastest_lap")

            tree.insert("", tk.END, iid=str(i), values=(d_name, f_pos, c_pos, pts, str(fl)))

        form_frame = tk.Frame(editor_win, bg="#090d16", pady=10)
        form_frame.pack(fill=tk.X, padx=15)

        tk.Label(form_frame, text="Ovr Pos:", bg="#090d16", fg="#94a3b8").grid(row=0, column=0, sticky="w")
        pos_entry = ttk.Entry(form_frame, width=6)
        pos_entry.grid(row=0, column=1, padx=4)

        tk.Label(form_frame, text="Cls Pos:", bg="#090d16", fg="#94a3b8").grid(row=0, column=2, sticky="w")
        cpos_entry = ttk.Entry(form_frame, width=6)
        cpos_entry.grid(row=0, column=3, padx=4)

        tk.Label(form_frame, text="Points:", bg="#090d16", fg="#94a3b8").grid(row=0, column=4, sticky="w")
        pts_entry = ttk.Entry(form_frame, width=6)
        pts_entry.grid(row=0, column=5, padx=4)

        tk.Label(form_frame, text="Fastest Lap:", bg="#090d16", fg="#94a3b8").grid(row=0, column=6, sticky="w")
        fl_combo = ttk.Combobox(form_frame, values=["true", "false"], width=7, state="readonly")
        fl_combo.grid(row=0, column=7, padx=4)

        def on_tree_select(event):
            sel = tree.selection()
            if not sel:
                return
            item_id = int(sel[0])
            res = results[item_id]
            f_pos = res.get("finish_position", 0)
            c_pos = res.get("class_finish_position", 0)
            pts = res.get("points_gain", 0.0)
            fl = res.get("fastest_lap", False)

            pos_entry.delete(0, tk.END); pos_entry.insert(0, str(f_pos))
            cpos_entry.delete(0, tk.END); cpos_entry.insert(0, str(c_pos))
            pts_entry.delete(0, tk.END); pts_entry.insert(0, str(pts))
            fl_combo.set(str(fl).lower())

        tree.bind("<<TreeviewSelect>>", on_tree_select)

        def apply_edit():
            sel = tree.selection()
            if not sel:
                return
            item_id = int(sel[0])
            try:
                new_pos = int(pos_entry.get())
                new_cpos = int(cpos_entry.get())
                new_pts = float(pts_entry.get())
                new_fl = fl_combo.get() == "true"
            except ValueError:
                messagebox.showerror("Error", "Invalid numeric values entered.")
                return

            res = results[item_id]
            res["finish_position"] = new_pos
            res["class_finish_position"] = new_cpos
            res["points_gain"] = new_pts
            res["fastest_lap"] = new_fl
            d_name = res["driver_name"]

            tree.item(str(item_id), values=(d_name, new_pos, new_cpos, new_pts, str(new_fl)))

        tk.Button(form_frame, text="Update Row", command=apply_edit, bg="#334155", fg="white", relief=tk.FLAT, font=("Helvetica", 9, "bold")).grid(row=0, column=8, padx=8)

        def save_and_close():
            overrides[r_key] = results
            self.override_mgr.save_overrides(champ_id, overrides)
            messagebox.showinfo("Success", "Round overrides saved successfully.")
            editor_win.destroy()
            self.refresh_current_championship_view()

        tk.Button(editor_win, text="Save & Persist Overrides", command=save_and_close, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(pady=10)

    def calculate_ti(self):
        races = self.active_profile['races']
        if not races: 
            return 0.5
        
        mode = getattr(self, 'ti_mode_var', None)
        mode_val = mode.get() if mode else "Advanced Ti"

        # 1. LEGACY TI SYSTEM (FIXED TRUE CHAMPIONSHIP LOOKUP)
        if "Legacy" in mode_val:
            total_score = 0.0
            for r in races:
                g_size = max(2, int(r.get('grid', 20)))
                f_pos = max(1, min(r['f'], g_size))
                c_pos = max(1, min(r['c'], g_size))
                
                score_champ = max(0.0, 1.0 - ((c_pos - 1) / float(g_size - 1 if g_size > 1 else 1)))
                score_race = max(0.0, 1.0 - ((f_pos - 1) / float(g_size - 1 if g_size > 1 else 1)))
                
                total_score += (0.70 * score_champ) + (0.30 * score_race)
            return max(0.0, min(1.0, total_score / len(races)))

        # 2. NEW ADVANCED NON-LINEAR TI SYSTEM WITH FIELD-PERCENTAGE RECOVERY & QUALIFYING WEIGHTING
        total_score = 0.0
        for r in races:
            g_size = max(2, int(r.get('grid', 20)))
            f_pos = max(1, min(r['f'], g_size))
            c_pos = max(1, min(r['c'], g_size))
            q_pos = max(1, min(int(r.get('q', 1)), g_size))
            
            # Non-linear normalization curves for Championship, Race Finish, and Qualifying
            champ_norm = 1.0 - ((c_pos - 1) / float(g_size - 1))
            score_champ = champ_norm ** 1.25

            race_norm = 1.0 - ((f_pos - 1) / float(g_size - 1))
            score_race = race_norm ** 1.25

            quali_norm = 1.0 - ((q_pos - 1) / float(g_size - 1))
            score_quali = quali_norm ** 1.1

            # Field-percentage scaled recovery delta (fraction of grid gained)
            positions_gained = q_pos - f_pos  # positive if you gained places
            field_fraction_gained = positions_gained / float(g_size - 1 if g_size > 1 else 1)

            # Graduated recovery reward curve scaled smoothly relative to field percentage
            if field_fraction_gained > 0.0:
                # Exponential reward boost for substantial field-percentage charges
                recovery_bonus = min(0.35, (field_fraction_gained ** 0.8) * 0.40)
            elif field_fraction_gained < 0.0:
                # Mild, controlled penalty for dropping positions from the start
                recovery_bonus = max(-0.15, field_fraction_gained * 0.15)
            else:
                recovery_bonus = 0.0  # Maintained position perfectly

            # Weighted blend: 40% Championship Standings, 35% Non-linear Race Finish, 25% Qualifying/Recovery
            base_weekend_score = (0.40 * score_champ) + (0.35 * score_race) + (0.25 * score_quali)
            final_race_score = base_weekend_score + recovery_bonus

            total_score += max(0.0, min(1.15, final_race_score))

        return max(0.0, min(1.0, total_score / len(races)))

    def refresh_dashboard(self):
        p = self.active_profile
        contract = p['currentContract']
        champ_name = contract.get('champName')
        
        champ_obj = next((c for c in engine.championships if c['name'].lower() == champ_name.lower()), {})
        start_type = champ_obj.get('start_type', 'Rolling')
        field_type = champ_obj.get('field_type', 'Non-Identical')
        parc_ferme = champ_obj.get('parc_ferme', 'False')

        categories = champ_obj.get('categories', [])
        classes_str = ", ".join(categories) if categories else contract.get('carClass', 'General')

        self.dash_header.config(text=f"{p['name']} | Season {p['season']} | {champ_name} ({contract.get('teamName')} - {contract.get('carModel')})")
        
        bp_disp = p.get('blueprint_name', '')
        bp_suffix = f"  |  Blueprint: {bp_disp}" if bp_disp else ""
        self.meta_lbl.config(text=f"Series Specs | Start: {start_type}  |  Field: {field_type}  |  Parc Fermé: {parc_ferme}  |  Prestige: {contract.get('prestige', 'N/A')}  |  Classes: {classes_str}{bp_suffix}")

        self.cal_listbox.delete(0, tk.END)
        cal = p.get('calendar', [])
        for idx, ev in enumerate(cal):
            self.cal_listbox.insert(tk.END, f"R{idx+1}: {ev}")

        for item in self.class_details_tree.get_children():
            self.class_details_tree.delete(item)

        active_cats = categories if categories else [contract.get('carClass', 'General')]
        for cat in active_cats:
            cat_clean = cat.strip()
            matching_models = sorted(list(set([car['car_model'] for car in engine.cars if car['category'].lower().strip() == cat_clean.lower()])))
            models_display = ", ".join(matching_models) if matching_models else "Spec / All Models"
            
            self.class_details_tree.insert("", tk.END, values=(cat_clean, models_display))

        race_rows = []
        for r in p['races']:
            race_rows.append((r['event'], f"P{r['q']}", f"P{r['f']}", f"P{r['c']}", r.get('grid', 20)))
        
        for i in self.race_tree.get_children(): self.race_tree.delete(i)
        for idx, row_vals in enumerate(race_rows):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.race_tree.insert("", tk.END, iid=str(idx), values=row_vals, tags=(tag,))
        self.race_tree.tag_configure("evenrow", background="#1e293b")
        self.race_tree.tag_configure("oddrow", background="#162032")

        ti = self.calculate_ti()
        if ti >= 0.75:
            self.rank_box.config(bg="#059669")
            self.rank_title_lbl.config(text="PROMOTION", bg="#059669")
            self.rank_sub_lbl.config(text=f"Trajectory: {ti*100:.1f}%", bg="#059669")
        elif ti >= 0.40:
            self.rank_box.config(bg="#d97706")
            self.rank_title_lbl.config(text="RETAINED", bg="#d97706")
            self.rank_sub_lbl.config(text=f"Trajectory: {ti*100:.1f}%", bg="#d97706")
        else:
            self.rank_box.config(bg="#dc2626")
            self.rank_title_lbl.config(text="DEMOTION", bg="#dc2626")
            self.rank_sub_lbl.config(text=f"Trajectory: {ti*100:.1f}%", bg="#dc2626")

        try:
            self.parsed_champs = self.parser.parse_championships()
            self.refresh_current_championship_view()
        except Exception:
            pass

    def delete_selected_race(self):
        sel = self.race_tree.selection()
        if not sel: return
        del self.active_profile['races'][int(sel[0])]
        self.save_active_profile()
        self.refresh_dashboard()

    def save_active_profile(self):
        p = self.active_profile
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("UPDATE careers SET season = ?, current_contract = ?, races = ?, calendar = ?, history = ?, blueprint_name = ?, second_monitor_id = ? WHERE id = ?",
                       (p['season'], json.dumps(p['currentContract']), json.dumps(p['races']), json.dumps(p['calendar']), json.dumps(p.get('history', [])), p.get('blueprint_name', ''), p.get('second_monitor_id', ''), p['id']))
        conn.commit()
        conn.close()

    # ------------------ SCREEN 3: SUMMARY & OFFERS ------------------
    def build_summary_screen(self):
        frame = tk.Frame(self.container, bg="#090d16", padx=25, pady=20)
        self.screens['summary'] = frame

        self.sum_title = tk.Label(frame, text="", font=("Helvetica", 14, "bold"), bg="#090d16", fg="#f8fafc")
        self.sum_title.pack(anchor="w")

        tk.Label(frame, text="Contract Market Offers (Full Continuum Range, Car Classes & Variable Seats)", font=("Helvetica", 10, "bold"), bg="#090d16", fg="#94a3b8").pack(anchor="w", pady=(6, 10))

        offer_container = tk.Frame(frame, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        offer_container.pack(fill=tk.BOTH, expand=True)

        self.offer_tree = ttk.Treeview(offer_container, columns=("tier_type", "series", "car_class", "team", "car", "tags", "prestige"), show="headings", height=12)
        self.offer_tree.heading("tier_type", text="Offer Type")
        self.offer_tree.heading("series", text="Championship Series")
        self.offer_tree.heading("car_class", text="Car Class")
        self.offer_tree.heading("team", text="Team Name")
        self.offer_tree.heading("car", text="Car Model")
        self.offer_tree.heading("tags", text="Tags")
        self.offer_tree.heading("prestige", text="Series Prestige")
        
        self.offer_tree.column("tier_type", width=150, stretch=False)
        self.offer_tree.column("series", width=200, stretch=False)
        self.offer_tree.column("car_class", width=100, stretch=False)
        self.offer_tree.column("team", width=180, stretch=False)
        self.offer_tree.column("car", width=180, stretch=False)
        self.offer_tree.column("tags", width=180, stretch=True)
        self.offer_tree.column("prestige", width=100, stretch=False)

        h_scroll = ttk.Scrollbar(offer_container, orient="horizontal", command=self.offer_tree.xview)
        v_scroll = ttk.Scrollbar(offer_container, orient="vertical", command=self.offer_tree.yview)
        self.offer_tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        self.offer_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        offer_container.grid_rowconfigure(0, weight=1)
        offer_container.grid_columnconfigure(0, weight=1)

        tk.Button(frame, text="Sign Selected Contract Offer", command=self.sign_contract, bg="#2563eb", fg="white", font=("Helvetica", 10, "bold"), pady=8, relief=tk.FLAT).pack(fill=tk.X, pady=15)

    def evaluate_season(self):
        self.sum_title.config(text=f"Season {self.active_profile['season']} Market Evaluation")
        for i in self.offer_tree.get_children(): self.offer_tree.delete(i)

        current_champ_name = self.active_profile['currentContract'].get('champName', '')
        current_champ = next((c for c in engine.championships if c['name'].lower() == current_champ_name.lower()), {"prestige": 5, "name": current_champ_name, "categories": [], "tags": []})
        current_prestige = current_champ['prestige']
        
        ti = self.calculate_ti()
        offers_pool = []

        sorted_champs = sorted(engine.championships, key=lambda x: x['prestige'])

        if ti >= 0.75:
            offers_pool.append(("Retention (Current Tier)", current_champ))

            max_prestige_jump = int(5 + ((ti - 0.75) / 0.25) * 10)
            target_max_prestige = current_prestige + max_prestige_jump

            promotion_candidates = [c for c in sorted_champs if current_prestige < c['prestige'] <= target_max_prestige]
            
            for c in promotion_candidates:
                offers_pool.append(("Promotion (Higher Tier)", c))
            
            lateral_champs = [c for c in sorted_champs if c['prestige'] == current_prestige and c['name'].lower() != current_champ_name.lower()]
            for c in lateral_champs[:2]:
                offers_pool.append(("Lateral (Same Tier)", c))

            fallback_champs = [c for c in sorted_champs if c['prestige'] < current_prestige]
            if fallback_champs:
                offers_pool.append(("Fallback (Lower Tier)", max(fallback_champs, key=lambda x: x['prestige'])))

        elif ti >= 0.40:
            retention_count = int(1 + (ti * 2))
            for _ in range(max(1, retention_count)):
                offers_pool.append(("Retention (Current Tier)", current_champ))

            lateral_champs = [c for c in sorted_champs if c['prestige'] == current_prestige and c['name'].lower() != current_champ_name.lower()]
            for c in lateral_champs[:2]:
                offers_pool.append(("Lateral (Same Tier)", c))

            fallback_champs = [c for c in sorted_champs if c['prestige'] < current_prestige]
            if fallback_champs:
                for c in sorted(fallback_champs, key=lambda x: x['prestige'], reverse=True)[:2]:
                    offers_pool.append(("Fallback (Lower Tier)", c))

        else:
            drop_depth = int(5 + ((1.0 - ti) * 20))
            fallback_champs = [c for c in sorted_champs if c['prestige'] < current_prestige and c['prestige'] >= current_prestige - drop_depth]
            if not fallback_champs:
                fallback_champs = [c for c in sorted_champs if c['prestige'] < current_prestige]

            num_demotions = int(2 + ((1.0 - ti) * 3))
            if fallback_champs:
                selected_demotions = random.sample(fallback_champs, min(num_demotions, len(fallback_champs)))
                for c in selected_demotions:
                    offers_pool.append(("Demotion / Fallback Offer", c))
            else:
                offers_pool.append(("Fallback Offer", current_champ))

        self.current_offers = []
        
        def get_strictly_matched_cars(ch):
            champ_cats = [cat.lower().strip() for cat in ch.get('categories', [])]
            champ_name_lower = ch['name'].lower()
            
            matched = []
            for car in engine.cars:
                car_cat = car['category'].lower().strip()
                # FIXED: Strict exact match to prevent shorthand category leaks
                if any(cc == car_cat for cc in champ_cats):
                    matched.append(car)
            
            if not matched:
                for car in engine.cars:
                    car_cat = car['category'].lower().strip()
                    if any(w in car_cat for w in champ_name_lower.split() if len(w) > 3):
                        matched.append(car)
                        
            return matched if matched else [{"team_name": "Factory Squad", "category": "Spec", "car_model": "Spec Race Car"}]

        offer_rows = []
        for type_label, ch in offers_pool:
            available_cars = get_strictly_matched_cars(ch)
            
            roll = random.random()
            if ti >= 0.85:
                base_seats = 4 if roll > 0.3 else 5
            elif ti >= 0.50:
                base_seats = 3 if roll > 0.4 else (4 if roll > 0.1 else 2)
            else:
                base_seats = 2 if roll > 0.4 else 3
                
            num_team_seats = min(base_seats, len(available_cars))
            num_team_seats = max(1, num_team_seats)
            
            chosen_cars = random.sample(available_cars, num_team_seats)
            tags_str = ", ".join(ch['tags']) if ch['tags'] else "Standard"
            for car_info in chosen_cars:
                car_class = car_info.get('category', 'General')
                offer = {
                    "champName": ch['name'], 
                    "teamName": car_info['team_name'], 
                    "carModel": car_info['car_model'], 
                    "carClass": car_class,
                    "prestige": ch['prestige']
                }
                self.current_offers.append(offer)
                offer_rows.append((type_label, ch['name'], car_class, car_info['team_name'], car_info['car_model'], tags_str, f"Prestige {ch['prestige']}"))

        for idx, row_vals in enumerate(offer_rows):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.offer_tree.insert("", tk.END, iid=str(idx), values=row_vals, tags=(tag,))
        self.offer_tree.tag_configure("evenrow", background="#1e293b")
        self.offer_tree.tag_configure("oddrow", background="#162032")

        self.show_screen('summary')

    def sign_contract(self):
        sel = self.offer_tree.selection()
        if not sel: return
        offer = self.current_offers[int(sel[0])]
        
        history_entry = {
            "season": self.active_profile['season'],
            "contract": self.active_profile['currentContract'],
            "ti": self.calculate_ti(),
            "races": self.active_profile['races']
        }
        if 'history' not in self.active_profile:
            self.active_profile['history'] = []
        self.active_profile['history'].append(history_entry)

        new_cal = engine.generate_calendar_for_champ(offer['champName'])

        self.active_profile['currentContract'] = offer
        self.active_profile['season'] += 1
        self.active_profile['races'] = []
        self.active_profile['calendar'] = new_cal
        self.active_profile['second_monitor_id'] = ''
        self.save_active_profile()

        messagebox.showinfo("Signed!", f"Signed with {offer['teamName']} for {offer['champName']} ({offer.get('carClass', 'General')})!")
        self.refresh_dashboard()
        self.show_screen('dashboard')

    # ------------------ SCREEN 4: GLOBAL HISTORY & SAVE INSPECTOR ------------------
    def build_history_screen(self):
        frame = tk.Frame(self.container, bg="#090d16", padx=25, pady=20)
        self.screens['history'] = frame

        header_row = tk.Frame(frame, bg="#090d16")
        header_row.pack(fill=tk.X, pady=(0, 10))

        tk.Label(header_row, text="Career Profile History & Championship Archive", font=("Helvetica", 14, "bold"), bg="#090d16", fg="#f8fafc").pack(side=tk.LEFT)

        sel_frame = tk.Frame(header_row, bg="#090d16")
        sel_frame.pack(side=tk.RIGHT)

        tk.Label(sel_frame, text="Select Career Profile:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#94a3b8").pack(side=tk.LEFT, padx=(0, 6))
        self.hist_profile_combo = ttk.Combobox(sel_frame, state="readonly", width=30)
        self.hist_profile_combo.pack(side=tk.LEFT)
        self.hist_profile_combo.bind("<<ComboboxSelected>>", self.on_history_profile_select)

        content = tk.Frame(frame, bg="#090d16")
        content.pack(fill=tk.BOTH, expand=True)

        left_col = tk.Frame(content, bg="#1e293b", width=340, highlightbackground="#334155", highlightthickness=1, padx=12, pady=12)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_col.pack_propagate(False)

        tk.Label(left_col, text="Completed Seasons", font=("Helvetica", 10, "bold"), bg="#1e293b", fg="#f8fafc").pack(anchor="w")

        self.hist_season_tree = ttk.Treeview(left_col, columns=("season", "series", "ti"), show="headings", height=12)
        self.hist_season_tree.heading("season", text="Season")
        self.hist_season_tree.heading("series", text="Championship")
        self.hist_season_tree.heading("ti", text="Score")
        
        self.hist_season_tree.column("season", width=70, stretch=False)
        self.hist_season_tree.column("series", width=160, stretch=False)
        self.hist_season_tree.column("ti", width=70, stretch=False)
        
        self.hist_season_tree.pack(fill=tk.BOTH, expand=True, pady=8)
        self.hist_season_tree.bind("<<TreeviewSelect>>", self.on_select_history_season)

        right_container = tk.Frame(content, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        right_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(right_container, text="Season Race Log Breakdown & Standings Archive", font=("Helvetica", 10, "bold"), bg="#1e293b", fg="#f8fafc", padx=12, pady=10).pack(anchor="w")
        
        self.hist_race_tree = ttk.Treeview(right_container, columns=("event", "q", "f", "c", "grid"), show="headings", height=10)
        self.hist_race_tree.heading("event", text="Track & Layout Variant")
        self.hist_race_tree.heading("q", text="Quali")
        self.hist_race_tree.heading("f", text="Finish (Class)")
        self.hist_race_tree.heading("c", text="Champ Pos")
        self.hist_race_tree.heading("grid", text="Grid Size")

        h_scroll_h = ttk.Scrollbar(right_container, orient="horizontal", command=self.hist_race_tree.xview)
        v_scroll_h = ttk.Scrollbar(right_container, orient="vertical", command=self.hist_race_tree.yview)
        self.hist_race_tree.configure(xscrollcommand=h_scroll_h.set, yscrollcommand=v_scroll_h.set)

        self.hist_race_tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

    def show_history_screen(self):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, driver_name FROM careers")
        self.history_profiles_cache = c.fetchall()
        conn.close()

        if not self.history_profiles_cache:
            messagebox.showinfo("No Careers", "No career profiles found in the database.")
            return

        combo_values = [f"{p[1]} ({p[2]})" for p in self.history_profiles_cache]
        self.hist_profile_combo['values'] = combo_values

        target_idx = 0
        if self.active_profile:
            for idx, p in enumerate(self.history_profiles_cache):
                if p[0] == self.active_profile['id']:
                    target_idx = idx
                    break

        self.hist_profile_combo.current(target_idx)
        self.load_history_for_profile_idx(target_idx)
        self.show_screen('history')

    def on_history_profile_select(self, event):
        idx = self.hist_profile_combo.current()
        if idx >= 0:
            self.load_history_for_profile_idx(idx)

    def load_history_for_profile_idx(self, idx):
        p_id = self.history_profiles_cache[idx][0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT history FROM careers WHERE id = ?", (p_id,))
        row = c.fetchone()
        conn.close()

        history = json.loads(row[0]) if row and row[0] else []
        
        for i in self.hist_season_tree.get_children(): self.hist_season_tree.delete(i)
        for i in self.hist_race_tree.get_children(): self.hist_race_tree.delete(i)

        season_rows = []
        for h in history:
            season_rows.append((f"Season {h['season']}", h['contract'].get('champName'), f"{h['ti']*100:.1f}%"))
            
        for s_idx, row_vals in enumerate(season_rows):
            tag = "evenrow" if s_idx % 2 == 0 else "oddrow"
            self.hist_season_tree.insert("", tk.END, iid=str(s_idx), values=row_vals, tags=(tag,))
        self.hist_season_tree.tag_configure("evenrow", background="#1e293b")
        self.hist_season_tree.tag_configure("oddrow", background="#162032")

        self.current_inspected_history = history

    def on_select_history_season(self, event):
        sel = self.hist_season_tree.selection()
        if not sel: return
        idx = int(sel[0])
        history = getattr(self, 'current_inspected_history', [])
        if idx >= len(history):
            return
        h = history[idx]
        
        hist_rows = []
        for r in h.get('races', []):
            hist_rows.append((r['event'], f"P{r['q']}", f"P{r['f']}", f"P{r['c']}", r.get('grid', 20)))

        for i in self.hist_race_tree.get_children(): self.hist_race_tree.delete(i)
        for r_idx, row_vals in enumerate(hist_rows):
            tag = "evenrow" if r_idx % 2 == 0 else "oddrow"
            self.hist_race_tree.insert("", tk.END, values=row_vals, tags=(tag,))
        self.hist_race_tree.tag_configure("evenrow", background="#1e293b")
        self.hist_race_tree.tag_configure("oddrow", background="#162032")

    # ------------------ PROFILE CREATION ------------------
    def open_new_career_dialog(self):
        if not engine.is_loaded:
            messagebox.showwarning("Load Blueprint First", "Please import and select an active Excel blueprint from the library first!")
            return

        dlg = tk.Toplevel(self)
        dlg.title("Start New Career Profile")
        dlg.geometry("780x540")
        dlg.configure(bg="#090d16")
        
        tk.Label(dlg, text="Career Profile Name:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(15, 2))
        ent_name = ttk.Entry(dlg, width=50); ent_name.pack(padx=20)

        tk.Label(dlg, text="Driver Name:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(8, 2))
        ent_driver = ttk.Entry(dlg, width=50); ent_driver.pack(padx=20)

        tk.Label(dlg, text=f"Active Source Blueprint: {engine.active_blueprint_filename}", font=("Helvetica", 8, "italic"), bg="#090d16", fg="#10b981").pack(anchor="w", padx=20, pady=(6, 2))
        tk.Label(dlg, text="Select Starting Championship:", font=("Helvetica", 9, "bold"), bg="#090d16", fg="#f8fafc").pack(anchor="w", padx=20, pady=(8, 4))

        table_frame = tk.Frame(dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

        champ_tree = ttk.Treeview(table_frame, columns=("name", "prestige", "categories", "tags"), show="headings", height=8)
        champ_tree.heading("name", text="Championship Series")
        champ_tree.heading("prestige", text="Prestige")
        champ_tree.heading("categories", text="Car Categories")
        champ_tree.heading("tags", text="Tags")
        
        champ_tree.column("name", width=250, stretch=False)
        champ_tree.column("prestige", width=85, stretch=False)
        champ_tree.column("categories", width=180, stretch=False)
        champ_tree.column("tags", width=200, stretch=True)

        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=champ_tree.xview)
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=champ_tree.yview)
        champ_tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        champ_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        sorted_champs = sorted(engine.championships, key=lambda x: x['prestige'])
        champ_rows = []
        for c in sorted_champs:
            cats_str = ", ".join(c['categories']) if c['categories'] else "General"
            tags_str = ", ".join(c['tags']) if c['tags'] else "Standard"
            champ_rows.append((c['name'], f"Prestige {c['prestige']}", cats_str, tags_str))

        for idx, row_vals in enumerate(champ_rows):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            champ_tree.insert("", tk.END, iid=str(idx), values=row_vals, tags=(tag,))
        champ_tree.tag_configure("evenrow", background="#1e293b")
        champ_tree.tag_configure("oddrow", background="#162032")

        if sorted_champs:
            champ_tree.selection_set("0")

        def prompt_car_selection(matching_cars, callback):
            car_dlg = tk.Toplevel(dlg)
            car_dlg.title("Select Team & Car Configuration")
            car_dlg.geometry("540x380")
            car_dlg.configure(bg="#090d16")
            car_dlg.transient(dlg)
            car_dlg.grab_set()

            tk.Label(car_dlg, text="Choose Your Starting Team & Vehicle", font=("Helvetica", 11, "bold"), bg="#090d16", fg="#f8fafc").pack(pady=(15, 5))
            tk.Label(car_dlg, text="Select your preferred team and car combination for this championship:", font=("Helvetica", 9), bg="#090d16", fg="#94a3b8").pack(pady=(0, 10))

            car_table_frame = tk.Frame(car_dlg, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
            car_table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

            car_tree = ttk.Treeview(car_table_frame, columns=("class", "team", "model"), show="headings", height=8)
            car_tree.heading("class", text="Car Class")
            car_tree.heading("team", text="Team Name")
            car_tree.heading("model", text="Car Model")

            car_tree.column("class", width=120, stretch=False)
            car_tree.column("team", width=190, stretch=False)
            car_tree.column("model", width=190, stretch=True)

            car_v_scroll = ttk.Scrollbar(car_table_frame, orient="vertical", command=car_tree.yview)
            car_tree.configure(yscrollcommand=car_v_scroll.set)

            car_tree.grid(row=0, column=0, sticky="nsew")
            car_v_scroll.grid(row=0, column=1, sticky="ns")
            car_table_frame.grid_rowconfigure(0, weight=1)
            car_table_frame.grid_columnconfigure(0, weight=1)

            car_rows = []
            for car in matching_cars:
                car_rows.append((car.get('category', 'General'), car['team_name'], car['car_model']))

            for idx, row_vals in enumerate(car_rows):
                tag = "evenrow" if idx % 2 == 0 else "oddrow"
                car_tree.insert("", tk.END, iid=str(idx), values=row_vals, tags=(tag,))
            car_tree.tag_configure("evenrow", background="#1e293b")
            car_tree.tag_configure("oddrow", background="#162032")

            if matching_cars:
                car_tree.selection_set("0")

            def confirm_car_choice():
                sel_car = car_tree.selection()
                if not sel_car:
                    messagebox.showwarning("Selection Required", "Please select a team and car combination.")
                    return
                chosen_car = matching_cars[int(sel_car[0])]
                car_dlg.destroy()
                callback(chosen_car)

            tk.Button(car_dlg, text="Confirm Selection & Start Career", command=confirm_car_choice, bg="#059669", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=20, pady=(0, 15))

        def create():
            c_name = ent_name.get()
            d_name = ent_driver.get()
            if not c_name or not d_name: 
                messagebox.showwarning("Missing Fields", "Please enter both a career profile name and driver name.")
                return
            
            sel = champ_tree.selection()
            if not sel:
                messagebox.showwarning("Selection Required", "Please select a starting championship from the table.")
                return

            selected_idx = int(sel[0])
            start_champ = sorted_champs[selected_idx]
            ch_name = start_champ['name']
            
            cal = engine.generate_calendar_for_champ(ch_name)
            
            champ_cats = [cat.lower().strip() for cat in start_champ.get('categories', [])]
            # FIXED: Strict exact match check for initial career creation car filtering
            matching_cars = [car for car in engine.cars if any(cc == car['category'].lower().strip() for cc in champ_cats)]
            if not matching_cars:
                matching_cars = [{"team_name": "Factory Squad", "category": "Spec", "car_model": "Spec Race Car"}]

            def finalize_career_creation(chosen_car):
                team_name = chosen_car['team_name']
                car_model = chosen_car['car_model']
                car_class = chosen_car.get('category', 'General')

                c_id = f"prof_{random.randint(10000, 99999)}"
                initial_contract = {
                    "champName": ch_name, 
                    "teamName": team_name, 
                    "carModel": car_model, 
                    "carClass": car_class, 
                    "prestige": start_champ['prestige']
                }
                
                conn = sqlite3.connect(DB_FILE)
                c_db = conn.cursor()
                c_db.execute("INSERT INTO careers VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                             (c_id, c_name, d_name, 1, json.dumps(initial_contract), "[]", json.dumps(cal), "[]", engine.active_blueprint_filename, ""))
                conn.commit()
                conn.close()

                dlg.destroy()
                self.load_profiles_list()

            prompt_car_selection(matching_cars, finalize_career_creation)

        tk.Button(dlg, text="Initialize Career", command=create, bg="#2563eb", fg="white", font=("Helvetica", 10, "bold"), pady=6, relief=tk.FLAT).pack(fill=tk.X, padx=20, pady=(0, 15))

    def load_selected_career(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a career profile to load.")
            return

        selected_profile_id = sel[0]
        
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT id, name, driver_name, season, current_contract, races, calendar, history, blueprint_name, second_monitor_id FROM careers WHERE id = ?", (selected_profile_id,))
        r = c.fetchone()
        conn.close()

        if r:
            bp_name = r[8]
            sec_mon_id = r[9] if len(r) > 9 and r[9] else ""
            if bp_name:
                bp_path = os.path.join(BLUEPRINTS_DIR, bp_name)
                if os.path.exists(bp_path):
                    engine.load_from_excel(bp_path)
                    self.bp_combo.set(bp_name)
                    self.update_blueprint_width_and_label(bp_name)
                    self.status_lbl.config(text=f"Blueprint Status: Active '{bp_name}'", fg="#10b981")
                else:
                    messagebox.showwarning("Blueprint Missing", f"The blueprint file '{bp_name}' associated with this career was not found in your library folder.\nUsing currently loaded blueprint instead.")

            self.active_profile = {
                "id": r[0], "name": r[1], "driverName": r[2], "season": r[3],
                "currentContract": json.loads(r[4]), "races": json.loads(r[5]), "calendar": json.loads(r[6]),
                "history": json.loads(r[7]) if len(r) > 7 and r[7] else [],
                "blueprint_name": bp_name,
                "second_monitor_id": sec_mon_id
            }
            self.refresh_dashboard()
            self.show_screen('dashboard')

if __name__ == "__main__":
    app = AMS2CareerApp()
    app.mainloop()

```

## 8. Build, Packaging & Compilation Instructions

### 8.1 Environment Setup

1. Ensure Python 3.10+ is installed and accessible via command line (`py` launcher).

2. Install external dependencies:

   ```
   py -m pip install --upgrade openpyxl pyinstaller
   
   ```

### 8.2 Binary Compilation Command

Run the following build command in the project root directory (with optional `--icon=app_icon.ico`):

```
py -m PyInstaller --noconsole --onefile --icon=app_icon.ico ams2_career_desktop.py

```

### 8.3 Output Locations

* Intermediate Build Objects: `./build/` (Safe to delete post-build)

* Final Standalone Executable: `./dist/ams2_career_desktop.exe`

## 9. Historical Lessons Learned & Architecture Post-Mortem

### 9.1 The Tkinter Taskbar Icon Bottleneck

* **Issue**: Compiling Tkinter apps into standalone Windows binaries frequently results in the application failing to render a custom `.ico` on the Windows 11 Taskbar, displaying the fallback Tcl/Tk "feather" icon instead.

* **Root Cause**: Tkinter initializes its top-level window handle (`HWND`) through an underlying C wrapper (`tcl86t.dll`). Windows 11 DWM registers the default `AppUserModelID` to the executable wrapper rather than dynamically injected icon resources.

* **Resolution**: Implemented explicit Win32 API calls (`ctypes.windll.user32.LoadImageW` and `SendMessageW`) alongside unique process model ID assignment (`SetCurrentProcessExplicitAppUserModelID`) to bind custom icon handles at runtime.

### 9.2 JSON Serialization Type Mismatch Post-Mortem (v2)

* **Issue**: Attempting to edit round results straight from live XML parser data (where results are native `RaceResultEntry` dataclass instances) caused a `TypeError: Object of type RaceResultEntry is not JSON serializable` upon saving overrides.

* **Resolution**: Implemented a defensive dictionary normalization mapping inside `open_round_editor()` to ensure all dataclass instances are converted into standard JSON-serializable dictionaries prior to persistence.

### 9.3 Substring Category Collision & Shorthand Leak Post-Mortem (v2.3)

* **Issue**: In multi-discipline blueprints containing short class identifiers (e.g., `"P1"` alongside `"LMP1"`, or `"GT"` alongside `"GT3"`), substring checks (`cc in car_cat`) caused cross-category contamination during initial career setup, market contract generation, and automated grid population.

* **Resolution**: Replaced all substring category comparisons across `BlueprintEditorWindow.auto_populate_teams_for_championship()`, `AMS2CareerApp.evaluate_season()` (`get_strictly_matched_cars`), and `AMS2CareerApp.open_new_career_dialog()` with strict normalized equality (`cc == car_cat`), ensuring 100% category isolation.

### 9.4 Toolbar Layout Compression in Select Screen (v2.3)

* **Issue**: Displaying a redundant `Active File: <filename>` status label inline inside `bp_bar` squeezed the right-aligned action buttons (`Blueprint Editor` and `Manage Library`) when long `.xlsx` filenames were selected.

* **Resolution**: Removed the redundant inline label (retaining the dedicated `self.status_lbl` indicator immediately below) and dynamically sized the library combobox width (`max(35, min(len(filename) + 5, 55))`) to guarantee full button visibility across all resolutions.