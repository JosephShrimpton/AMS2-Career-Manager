# AMS2 Career Manager

A lightweight, robust, and standalone desktop application designed to bring structured career progression, dynamic contract market evaluations, telemetry-integrated standings, multi-class championship management, and an in-app blueprint editor to **Automobilista 2 (AMS2)**.

## Version 2.3 (Advanced TI System, In-App Blueprint Editor & Second Monitor XML Integration)

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/c4a5cda6-e366-4fa3-94b1-b2293f052bb8" />

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/0c63d488-1ba1-4906-a173-38275b8b277e" />

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/e33f6fc5-6f86-4c06-82c1-f35ef724e399" />

Version 2.3 transforms AMS2 Career Manager into an exhaustive, automated, telemetry-integrated career engine and blueprint authoring suite. While retaining complete manual fallback entry, v2.3 connects directly with **Second Monitor** XML data exports for live standings synchronization, automated qualifying position detection, dynamic multi-class scoring calculations, persistent round overrides, and algorithmic career contract offers via an upgraded **Advanced Non-Linear Trajectory Index (TI)** engine, alongside a built-in **Blueprint Creator & Editor**.

### New Key Features in v2.3

* **Advanced Non-Linear Trajectory Index (TI) & Recovery Engine**: Evaluates driver performance across seasons using non-linear normalization curves ($1.25$ exponent for championship standings and race finishes, $1.1$ for qualifying) combined with a graduated field-percentage recovery bonus/penalty ($\text{Positions Gained} / \text{Grid Size}$):
  
  $$\text{Score}_{\text{champ}} = \left(1.0 - \frac{\text{ChampPos} - 1}{\text{GridSize} - 1}\right)^{1.25}, \quad \text{Score}_{\text{race}} = \left(1.0 - \frac{\text{FinishPos} - 1}{\text{GridSize} - 1}\right)^{1.25}, \quad \text{Score}_{\text{quali}} = \left(1.0 - \frac{\text{QualiPos} - 1}{\text{GridSize} - 1}\right)^{1.1}$$
  
  $$\text{Weekend Score} = 0.40 \times \text{Score}_{\text{champ}} + 0.35 \times \text{Score}_{\text{race}} + 0.25 \times \text{Score}_{\text{quali}} + \text{Recovery Bonus}$$

  * **UI Ti System Toggle**: Switch seamlessly on the dashboard between *Advanced Ti (Non-Linear + Recovery)* and *Legacy Ti (Fixed)* evaluation modes.

* **In-App Blueprint Creator & Editor**: Build, stage, modify, and export custom multi-sheet Excel `.xlsx` blueprint workbooks directly inside the application without external spreadsheet software. Features inline double-click editing for championships, custom track calendar builders, category managers, and automated population of 40 authentic real-team car pairings per new championship.

* **Strict Category Matching Engine**: Enforces exact string matching during contract market generation, blueprint auto-population, and initial career setup to eliminate shorthand category leaks (e.g., preventing tags like "P1" from matching "LMP1").

* **Save Blueprint Re-assignment & Content Summary**: Re-bind or update the source `.xlsx` blueprint file for any existing career save on the fly, complete with an interactive content summary viewer displaying loaded championships, prestiges, and categories.

* **Direct Second Monitor XML Integration**: Live parsing of `Championships.xml` exports from Second Monitor to track real-time series standings and round classifications.

* **Independent Multi-Class Standings & Scope Filtering**: Dynamically calculates independent class standings, positions, and points allocation from combined scratch race leaderboards using XML scoring matrices, complete with UI scope filters for both Overall and Class-specific views.

* **Auto-Qualifying & Session Detection**: Automatically detects qualifying positions, driver identification, and class-specific field sizes directly from race session telemetry.

* **Persistent Round XML Override Manager**: Modify finishing positions, class positions, points gains, and fastest lap bonuses on a per-round basis without corrupting telemetry data or application state.

* **JSON Career Save Backup & Restore**: Export individual career profiles (including history, active contracts, calendars, and linked XML IDs) to portable `.json` files and import them into any installation.

* **Multi-Sheet Excel Blueprint Loader**: Parses comprehensive multi-sheet workbooks (`championships`, `events`, `cars`, `categories`) using `openpyxl`.

* **Global Multi-Season History & Save Inspector**: Complete multi-season archive viewer to inspect past career profiles, season scores, and detailed historical race breakdowns.

### How to Set Up & Use Version 2.3

#### Step 1: Install & Launch Second Monitor

1. Ensure **Second Monitor** (the community telemetry and timing application for AMS2) is installed on your PC.

2. Run Second Monitor while playing Automobilista 2 to ensure race telemetry and championship progress are logged to its default data path:
   `C:\Users\<Your-Username>\Documents\SecondMonitor\Championships\Championships.xml`

#### Step 2: Import or Create Blueprints

1. Launch **AMS2 Career Manager v2.3**.

2. In the **Blueprint Library** bar on the main screen:
   * Click **Import Blueprint (.xlsx)** to load an existing custom workbook into your AppData library (`%APPDATA%\AMS2CareerManager\blueprints\`).
   * Or click **Blueprint Editor** to visually create, edit, and export a brand-new `.xlsx` blueprint workbook from scratch.

3. Select your active blueprint from the **Blueprint Library** dropdown and click **+ New Career Profile** to choose your starting championship, team, and car.

#### Step 3: Bind Second Monitor XML in AMS2 Career Manager

1. Load your Career Profile to open the main Dashboard.

2. In the main Dashboard, navigate to the **Live XML Inspector** bar on the right column.

3. If your XML path differs from the default location, click **Set XML Path** and select your `Championships.xml` file.

4. Click **Reload XML** to pull the latest championship data into the manager.

#### Step 4: Link Your Active Career Championship Series

1. On the Dashboard, click **Link Series**.

2. A window will display all active championships detected inside Second Monitor's XML export.

3. Select the championship corresponding to your active career season and click **Confirm & Bind Link**.

4. The application will now lock onto that series, displaying live overall and independent class standings alongside session classifications.

#### Step 5: Lock In & Record Round Results

1. Go to the **Per-Round Classification & Result Recorder** tab.

2. Select the completed event from the **Select Round** drop-down (and optionally filter by **Round Scope** for multi-class races).

3. The app will auto-populate your detected qualifying position and class/overall grid size.

4. Click **Lock In & Record Round** to commit the round results toward your season's Trajectory Index (TI).

#### Step 6: (Optional) Override Round Telemetry, Edit Calendar Tracks, or Update Blueprint

* **Edit Round XML**: Select any round in the inspector tab and click **Edit Round XML** to adjust finishing positions, class positions, points gain, or fastest lap indicators. Overrides are saved safely to `%APPDATA%\AMS2CareerManager\round_overrides\`.
* **Edit Season Calendar**: Select a round in the **Season Event Calendar** on the left panel and click **Edit Track** to customize the track or layout variant for that season.
* **Update Save Blueprint**: Click **Update Save Blueprint** in the top metadata bar to re-bind a new or updated `.xlsx` blueprint to your active career save.

#### Step 7: End Season & Dynamic Contract Market

1. Once all season calendar events are complete, click **End Season & Evaluate Offers**.

2. Based on your cumulative Trajectory Index (TI) (evaluated via either *Advanced Ti* or *Legacy Ti*), review generated contract offers (Promotion, Retention, Lateral, Demotion) strictly matched to vehicle categories and teams from your blueprint.

3. Select your contract and click **Sign Selected Contract Offer** to advance to the next season and automatically archive your past season into the Global History log.

## Version 1.0 (Manual Entry Career Tracker)

*Note: Version 1.0 remains fully available on the Releases page for users who prefer completely manual entry without telemetry utilities.*

Built with Python and Tkinter, this tool acts as a dedicated career engine that loads external Excel race blueprints, logs contract earnings, tracks multi-season standings, and securely persists your career save states.

### V1 Key Features

<img width="1262" height="812" alt="image" src="https://github.com/user-attachments/assets/8cc9f09f-6841-48f1-9704-d4eccc38d431" />

<img width="1262" height="812" alt="image" src="https://github.com/user-attachments/assets/13da6c00-eceb-4fa1-865a-65c738b1d411" />

* **Slate Dark UI**: Clean, eye-friendly custom theme optimized for extended gaming sessions.

* **Persistent AppData Architecture**: Automatically manages databases and user preferences inside standard Windows application folders (`%APPDATA%\AMS2CareerManager\`), keeping your environment clean and portable.

* **Excel Blueprint Loader**: Import race calendars, championship structures, and scoring tables directly from `.xlsx` files.

* **Dynamic Table Controls**: Native zebra-striping, auto-scrolling, multi-column sorting, and real-time state updates.

* **Built-in Safety Locks**: Prevents accidental data overwrites and ensures clean database transactions.

* **Standalone Deployment**: Fully self-contained application with zero external runtime dependencies required for pre-built binaries.

### How to Use Version 1.0

1. **Launch the Application**: On startup, the manager automatically initializes or loads your active profile from `%APPDATA%\AMS2CareerManager\ams2_careers.db`.

2. **Import Blueprints**: Use the *Import Excel Blueprint* button to load customized championship layouts, prize money payouts, and race schedules.

3. **Log Race Results**: Enter your finishing positions, bonuses, and penalties into the session tracker manually.

4. **Track Progression**: Review multi-season statistics, driver standings, and cumulative career earnings directly within the UI.

## Getting Started & Installation

### Option 1: Downloading Pre-Built Executable (Recommended)

1. Go to the **Releases** section on the right side of this GitHub repository.

2. Select your desired version:

   * **v2.3 (Latest)**: Advanced Non-Linear TI & Recovery engine, In-App Blueprint Creator & Editor, Second Monitor XML integration, automated telemetry parsing, persistent override editor, dynamic market evaluation.

   * **v1.0 (Legacy)**: Purely manual entry career progression engine.

3. Download the corresponding `.exe` binary.

4. Run the executable directly—no installation or external dependencies required.

### Option 2: Running from Source

#### Prerequisites

* **Python 3.10+**

* Required packages: `openpyxl` (for Excel imports and blueprint editing)

#### Quickstart

1. Clone this repository:

   ```bash
   git clone https://github.com/your-username/ams2-career-manager.git
   cd ams2-career-manager
   ```

2. Install dependencies:

   ```bash
   pip install openpyxl
   ```

3. Run the application:

   ```bash
   python ams2_career_desktop.py
   ```

### Option 3: Compiling Your Own Portable Executable

If you want to build the executable manually from source:

1. Install PyInstaller:

   ```bash
   pip install pyinstaller openpyxl
   ```

2. Build the binary using the custom icon asset and window flags:

   ```bash
   pyinstaller --noconsole --onefile --icon=app_icon.ico ams2_career_desktop.py
   ```

3. Your compiled executable will be generated in the `dist/` directory.

## Repository Structure

```text
├── ams2_career_desktop.py  # Main application logic, XML parser, Blueprint Editor, TI engine & Tkinter UI
├── app_icon.ico             # Application icon asset
├── README.md                # Project documentation
└── .gitignore               # Build, cached, and local file exclusions
```

## Legal Disclaimer & Notice

**Automobilista 2** is a registered trademark of **Reiza Studios**.

* **Non-Affiliation**: This software is an independent, open-source community application. It is **not** affiliated with, endorsed by, sponsored by, or created by Reiza Studios or any of its partners.

* **Trademarks**: All brand names, game names, logos, and associated trademarks referenced within this application or documentation belong solely to their respective owners.

## License

This project is licensed under the **MIT License**.

```text
MIT License

Copyright (c) 2026

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```