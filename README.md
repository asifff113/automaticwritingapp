# Automatic Writing Assistant

Desktop app that lets a user paste text, then automatically type it into the currently focused field after a countdown.

## What it does

- Accepts pasted text in the app window.
- Starts typing after a configurable countdown so you can switch to your target window.
- Supports configurable typing speed.
- Lets you stop typing via:
  - `Stop` button in the app
  - `F9` emergency stop while typing (when global hotkeys are available)

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

1. Paste your text into the app.
2. Set:
   - `Start countdown (sec)` (example: `5`)
   - `Delay per character (sec)` (example: `0.03`)
3. Click `Start Typing`.
4. Immediately focus the input field where you want text to be typed.
5. Press `F9` to stop if available, otherwise use the app `Stop` button.

## Important

Use this only where automation is allowed by the platform or organization policy.
