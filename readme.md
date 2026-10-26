# AMS2 Career Manager Desktop

A lightweight, robust, and standalone desktop application designed to bring structured career progression, dynamic financial tracking, and multi-class championship management to Automobilista 2 (AMS2).

Built with Python and Tkinter, this tool acts as a dedicated career engine that loads external Excel race blueprints, logs contract earnings, tracks multi-season standings, and securely persists your career save states.

<img width="1262" height="812" alt="image" src="https://github.com/user-attachments/assets/e6458534-65b1-47f8-b688-b7f0ee01e717" />

<img width="1262" height="812" alt="image" src="https://github.com/user-attachments/assets/7f9afd45-6f95-4bc8-ba7a-549e676298f1" />

## Key Features

* **Slate Dark UI**: Clean, eye-friendly custom theme optimized for extended gaming sessions.
* **Persistent AppData Architecture**: Automatically manages databases and user preferences inside standard Windows application folders (`%APPDATA%\AMS2CareerManager\`), keeping your environment clean and portable.
* **Excel Blueprint Loader**: Import race calendars, championship structures, and scoring tables directly from `.xlsx` files.
* **Dynamic Table Controls**: Native zebra-striping, auto-scrolling, multi-column sorting, and real-time state updates.
* **Built-in Safety Locks**: Prevents accidental data overwrites and ensures clean database transactions.
* **Standalone Deployment**: Fully self-contained application with zero external runtime dependencies required for pre-built binaries.

## Getting Started

### Option 1: Downloading Pre-Built Executable (Recommended)

1. Go to the **Releases** section on the right side of this GitHub repository (or under the `Builds/Releases` tab).
2. Download the latest `ams2_career_desktop.exe` binary.
3. Run `ams2_career_desktop.exe` directly. No installation or external dependencies are required.

### Option 2: Running from Source

#### Prerequisites

* **Python 3.10+**
* Required packages: `openpyxl` (for Excel imports)

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

### Option 3: Compiling your own Portable Executable

If you want to build the executable manually from source:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Build the binary using the custom icon asset:
   ```bash
   pyinstaller --noconsole --onefile --icon=app_icon.ico ams2_career_desktop.py
   ```

3. Your compiled executable will be generated in the `dist/` directory.

## How to Use

1. **Launch the Application**: On startup, the manager automatically initializes or loads your active profile from `%APPDATA%\AMS2CareerManager\ams2_career.db`.
2. **Import Blueprints**: Use the *Import Excel Blueprint* button to load customized championship layouts, prize money payouts, and race schedules.
3. **Log Race Results**: Enter your finishing positions, bonuses, and penalties into the session tracker.
4. **Track Progression**: Review multi-season statistics, driver standings, and cumulative career earnings directly within the UI.

## Repository Structure

```text
├── ams2_career_desktop.py  # Main application logic & Tkinter UI implementation
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
