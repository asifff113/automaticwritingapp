# Automatic Writing Assistant v3.0

A beautiful desktop app that lets you paste text and automatically type it into any focused input field. Built with pure Python/tkinter - zero extra dependencies on Windows.

**[Download from the official website](https://asifff113.github.io/automaticwritingapp/)** | [GitHub Releases](https://github.com/asifff113/automaticwritingapp/releases)

---

## Features

### Core
- Paste or type text, then simulate natural keyboard input into any focused field
- Configurable countdown timer (1-30s) to switch to target window
- Adjustable typing speed (5-300ms per character)
- Emergency stop via F9 hotkey or Stop button
- Pause / Resume typing mid-session
- Live WPM (words per minute) display
- Progress bar with percentage and ETA

### Typing Modes
- **Constant** - Fixed delay between characters
- **Human-like** - Random speed variation with longer pauses at punctuation (`.!?,;:` and newlines)
- **Burst** - Type in fast bursts of 3-8 characters with pauses in between
- Configurable randomness slider (0-100%) for natural feel

### Speed Presets (New in v3.0)
- **Slow** (80ms) - Careful, deliberate typing
- **Normal** (30ms) - Balanced everyday speed
- **Fast** (15ms) - Quick typing
- **Blazing** (5ms) - Maximum speed

### Editor
- Full text editor with undo/redo, cut/copy/paste
- Right-click context menu
- Open text from `.txt` files (Ctrl+O)
- Save text to files (Ctrl+S)
- Paste from clipboard button
- Live character / word / line count
- Estimated typing time display
- Zoom in/out (Ctrl+Plus / Ctrl+Minus)
- Word wrap toggle

### Find & Replace (New in v3.0)
- Built-in Find & Replace bar (Ctrl+F)
- Navigate matches with Next/Previous
- Replace one or replace all
- Match highlighting

### Text Transforms (New in v3.0)
12 one-click text transforms:
- UPPERCASE, lowercase, Title Case, Sentence case
- Sort A-Z, Sort Z-A, Reverse lines
- Remove duplicates, Number lines
- Remove empty lines, Trim whitespace, Squeeze blank lines

### Statistics Dashboard (New in v3.0)
- Lifetime stats: total characters, sessions, time spent
- Session history with per-session details
- Persistent tracking across app restarts

### Presets
- Save frequently-used text as named presets
- Load, overwrite, and delete presets
- Presets stored in `presets.json` alongside the app

### Auto-Draft (New in v3.0)
- Automatically saves editor content between sessions
- Never lose your work - text is restored on next launch

### Recent Files (New in v3.0)
- Quick access menu for recently opened files

### Options
- Auto-minimize window before typing starts
- Restore window after typing completes
- Sound notification on completion (Windows)
- Skip empty lines
- Trim trailing whitespace per line
- Repeat typing 1-99 times
- Pin window always-on-top toggle
- Shift+Enter mode for chat apps (ChatGPT, Discord) - configurable

### Themes (New in v3.0)
- Light theme (default) and Dark theme
- One-click toggle in the header
- Persistent preference saved to disk

### UI
- 5 tabs: Editor, Settings, Live Log, Statistics, How to Use
- Refined color palette with hero-styled header
- Keyboard shortcuts: Ctrl+O (open), Ctrl+S (save), Ctrl+Enter (start), Ctrl+F (find)
- Export log to file
- Cross-platform: Windows, macOS, Linux

## Requirements

- Python 3.10+
- Desktop OS:
  - Windows (native `SendInput` backend - full emoji support)
  - macOS (via `pynput`)
  - Linux (via `pynput`)

## OS permissions

- macOS: grant Accessibility permission to the app/terminal so keyboard automation works.
- Linux: run inside a desktop session (X11 or Wayland) with input access.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

## Build portable executable

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
2. Upload the generated files from `dist/` to GitHub Releases.
3. Users download for their OS:
   - **Windows**: Direct `.exe` download - just run it
   - **macOS**: `.zip` containing the binary - extract and run
   - **Linux**: `.zip` containing the binary - extract, `chmod +x`, and run

## GitHub automated release builds

Workflow file: `.github/workflows/build-and-release.yml`

- Manual run: GitHub Actions -> `build-and-release` -> `Run workflow`
- Automatic release publish:
  1. Create and push a tag like `v3.0.0`
  2. Workflow builds all OS artifacts
  3. Workflow publishes a GitHub Release with attached files

## Device compatibility

- Windows/macOS/Linux desktop: supported
- Android/iOS: not supported by this desktop architecture

## Usage

1. Launch the app: `python app.py`
2. **Editor tab**: Paste or type your text, or open a `.txt` file
3. **Settings tab**: Choose typing mode, speed preset, countdown, and options
4. Click **Start Typing** (or press Ctrl+Enter)
5. Quickly switch to the target input field during the countdown
6. Watch the **Live Log** tab for real-time progress
7. Press **F9** to emergency-stop, or use **Pause** / **Stop** buttons
8. Check the **Statistics** tab for session history and lifetime stats

## Important

Use this only where automation is allowed by the platform or organization policy.
