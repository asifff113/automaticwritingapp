# Automatic Writing Assistant

A beautiful desktop app that lets you paste text and automatically type it into any focused input field after a countdown. Built with pure Python/tkinter - zero extra dependencies on Windows.

**[Download from the official website](https://asifff113.github.io/automaticwritingapp/)** | [GitHub Releases](https://github.com/asifff113/automaticwritingapp/releases)

## Features

**Core**
- Paste or type text, then simulate natural keyboard input into any focused field
- Configurable countdown timer (1-30s) to switch to target window
- Adjustable typing speed (5-300ms per character)
- Emergency stop via F9 hotkey or Stop button
- Pause / Resume typing mid-session

**Typing Modes**
- **Constant** - Fixed delay between characters
- **Human-like** - Random speed variation with longer pauses at punctuation (`.!?,;:` and newlines)
- **Burst** - Type in fast bursts of 3-8 characters with pauses in between
- Configurable randomness slider (0-100%) for natural feel

**Editor**
- Full text editor with undo/redo, cut/copy/paste
- Right-click context menu
- Open text from `.txt` files (Ctrl+O)
- Save text to files (Ctrl+S)
- Paste from clipboard button
- Live character / word / line count
- Estimated typing time display

**Presets**
- Save frequently-used text as named presets
- Load, overwrite, and delete presets
- Presets stored in `presets.json` alongside the app

**Options**
- Auto-minimize window before typing starts
- Restore window after typing completes
- Sound notification on completion (Windows)
- Skip empty lines
- Trim trailing whitespace per line
- Repeat typing 1-99 times
- Pin window always-on-top toggle

**Live Log**
- Real-time color-coded typing log with timestamps
- Shows mode, speed, progress, and errors
- Elapsed time and ETA displayed during typing

**UI**
- Dark themed interface with tabbed layout (Editor / Settings / Live Log)
- Keyboard shortcuts: Ctrl+O (open), Ctrl+S (save), Ctrl+Enter (start)
- Progress bar with percentage display
- Cross-platform: Windows, macOS, Linux

## Requirements

- Python 3.10+
- Desktop OS:
  - Windows (native `SendInput` backend)
  - macOS (via `pynput`)
  - Linux (via `pynput`)

## OS permissions

- macOS: grant Accessibility permission to the app/terminal so keyboard automation works.
- Linux: run inside a desktop session (not headless) with input access.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

## Build downloadable software

All builds produce:

- standalone executable/binary in `dist/`
- zip archive in `dist/AutomaticWritingAssistant-<os>.zip`

### Windows

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1 -Clean
```

### macOS

```bash
chmod +x ./build_macos.sh
./build_macos.sh --clean
```

### Linux

```bash
chmod +x ./build_linux.sh
./build_linux.sh --clean
```

### Generic (any desktop OS)

```bash
python -m pip install -r requirements-build.txt
python -m pip install -r requirements.txt
python build_release.py --clean
```

## Share with users

1. Build on each target OS (Windows/macOS/Linux).
2. Upload the generated zip files from `dist/` to GitHub Releases or cloud storage.
3. Users download the zip for their OS, extract, and run the executable.

## GitHub automated release builds

Workflow file: `.github/workflows/build-and-release.yml`

- Manual run: GitHub Actions -> `build-and-release` -> `Run workflow`
- Automatic release publish:
  1. Create and push a tag like `v1.0.0`
  2. Workflow builds all OS artifacts
  3. Workflow publishes a GitHub Release with attached zip files

## Device compatibility

- Windows/macOS/Linux desktop: supported.
- Android/iOS: not supported by this desktop architecture.

## Usage

1. Launch the app: `python app.py`
2. **Editor tab**: Paste or type your text, or open a `.txt` file
3. **Settings tab**: Choose typing mode, speed, countdown, and options
4. Click **Start Typing** (or press Ctrl+Enter)
5. Quickly switch to the target input field during the countdown
6. Watch the **Live Log** tab for real-time progress
7. Press **F9** to emergency-stop, or use **Pause** / **Stop** buttons

## Important

Use this only where automation is allowed by the platform or organization policy.
