# AMS2 Career Manager

A lightweight, robust, and standalone desktop application designed to bring structured career progression, dynamic contract market evaluations, telemetry-integrated standings, and multi-class championship management to **Automobilista 2 (AMS2)**.

## Version 2.0 (Telemetry & Second Monitor XML Integration)

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/c4a5cda6-e366-4fa3-94b1-b2293f052bb8" />

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/0c63d488-1ba1-4906-a173-38275b8b277e" />

<img width="1282" height="852" alt="image" src="https://github.com/user-attachments/assets/e33f6fc5-6f86-4c06-82c1-f35ef724e399" />

Version 2.0 transforms AMS2 Career Manager into an automated, telemetry-integrated career engine. While retaining complete manual fallback entry, v2 connects directly with **Second Monitor** XML data exports for live standings synchronization, automated qualifying position detection, dynamic multi-class scoring calculations, persistent round overrides, and algorithmic career contract offers via the Trajectory Index (TI).

### New Key Features in v2.0

* **Direct Second Monitor XML Integration**: Live parsing of `Championships.xml` exports from Second Monitor to track real-time series standings and round classifications.

* **Independent Multi-Class Standings Engine**: Dynamically calculates independent class standings, positions, and points allocation from combined scratch race leaderboards.

* **Auto-Qualifying & Session Detection**: Automatically detects qualifying positions, driver identification, and field sizes directly from race session telemetry.

* **Persistent Round XML Override Manager**: Modify finishing positions, points gains, and fastest lap bonuses on a per-round basis without corrupting telemetry data or application state.

* **Trajectory Index (TI) & Dynamic Contract Market**: Evaluates driver performance across seasons using a weighted performance formula:
  

  $$
  TI = 0.70 \times \left(1.0 - \frac{\text{ChampPos} - 1}{\text{GridSize} - 1}\right) + 0.30 \times \left(1.0 - \frac{\text{FinishPos} - 1}{\text{GridSize} - 1}\right)
  $$

  
  Generates dynamic Promotion, Retention/Lateral, and Demotion contract offers paired strictly to vehicle categories and teams from your blueprint.

* **Multi-Sheet Excel Blueprint Loader**: Parses comprehensive multi-sheet workbooks (`championships`, `events`, `cars`) using `openpyxl`.

* **Global Multi-Season History & Save Inspector**: Complete multi-season archive viewer to inspect past career profiles, season scores, and detailed historical race breakdowns.

### How to Set Up & Use Version 2.0 (Second Monitor Integration)

#### Step 1: Install & Launch Second Monitor

1. Ensure **Second Monitor** (the community telemetry and timing application for AMS2) is installed on your PC.

2. Run Second Monitor while playing Automobilista 2 to ensure race telemetry and championship progress are logged to its default data path:
   `C:\Users\<Your-Username>\Documents\SecondMonitor\Championships\Championships.xml`

#### Step 2: Bind Second Monitor XML in AMS2 Career Manager

1. Launch **AMS2 Career Manager v2.0**.

2. Select or create your Career Profile.

3. In the main Dashboard, navigate to the **Live XML Inspector** bar on the right column.

4. If your XML path differs from the default location, click **Set XML Path** and select your `Championships.xml` file.

5. Click **Reload XML** to pull the latest championship data into the manager.

#### Step 3: Link Your Active Career Championship Series

1. On the Dashboard, click **Link Series**.

2. A window will display all active championships detected inside Second Monitor's XML export.

3. Select the championship corresponding to your active career season and click **Confirm & Bind Link**.

4. The application will now lock onto that series, displaying live class standings and session classifications.

#### Step 4: Lock In & Record Round Results

1. Go to the **Per-Round Classification & Result Recorder** tab.

2. Select the completed event from the **Select Round** drop-down.

3. The app will auto-populate your detected qualifying position and grid size.

4. Click **Lock In & Record Round** to commit the round results toward your season's Trajectory Index (TI).

#### Step 5: (Optional) Override or Edit Round Telemetry

If Second Monitor recorded an error or you wish to adjust race results manually:

1. Select the round in the inspector tab and click **Edit Round XML**.

2. Adjust finishing position, class position, points gain, or fastest lap indicators for any driver on the grid.

3. Click **Save & Persist Overrides**. Overrides are saved safely to `%APPDATA%\AMS2CareerManager\round_overrides\`.

#### Step 6: End Season & Dynamic Contract Market

1. Once all season calendar events are complete, click **End Season & Evaluate Offers**.

2. Based on your cumulative Trajectory Index (TI), review generated contract offers (Promotion, Retention, Lateral, Demotion).

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

   * **v2.0 (Latest)**: Second Monitor XML integration, automated telemetry parsing, persistent override editor, dynamic market evaluation.

   * **v1.0 (Legacy)**: Purely manual entry career progression engine.

3. Download the corresponding `.exe` binary.

4. Run the executable directly—no installation or external dependencies required.

### Option 2: Running from Source

#### Prerequisites

* **Python 3.10+**

* Required packages: `openpyxl` (for Excel imports)

#### Quickstart

1. Clone this repository:

   ```
   git clone https://github.com/your-username/ams2-career-manager.git
   cd ams2-career-manager
   
   ```

2. Install dependencies:

   ```
   pip install openpyxl
   
   ```

3. Run the application:

   ```
   python ams2_career_desktop.py
   
   ```

### Option 3: Compiling Your Own Portable Executable

If you want to build the executable manually from source:

1. Install PyInstaller:

   ```
   pip install pyinstaller openpyxl
   
   ```

2. Build the binary using the custom icon asset and window flags:

   ```
   pyinstaller --noconsole --onefile --icon=app_icon.ico ams2_career_desktop.py
   
   ```

3. Your compiled executable will be generated in the `dist/` directory.

## Repository Structure

```
├── ams2_career_desktop.py  # Main application logic, XML parser, TI engine & Tkinter UI
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

```
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
