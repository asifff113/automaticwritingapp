import argparse
import os
import platform
import shutil
import subprocess
import sys
import zipfile

APP_NAME = "AutomaticWritingAssistant"
OS_LABELS = {
    "Windows": "windows",
    "Darwin": "macos",
    "Linux": "linux",
}


def remove_path(path):
    if os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.remove(path)


def main():
    parser = argparse.ArgumentParser(description="Build release artifacts with PyInstaller.")
    parser.add_argument("--clean", action="store_true", help="Remove prior build artifacts")
    args = parser.parse_args()

    system = platform.system()
    if system not in OS_LABELS:
        raise RuntimeError(f"Unsupported operating system: {system}")

    if args.clean:
        remove_path("build")
        remove_path("dist")
        remove_path(f"{APP_NAME}.spec")

    root_dir = os.path.dirname(os.path.abspath(__file__))
    icon_ico = os.path.join(root_dir, "icon.ico")
    icon_png = os.path.join(root_dir, "icon.png")

    # Generate icons if missing
    if not os.path.exists(icon_ico) or not os.path.exists(icon_png):
        try:
            from generate_icon import generate_ico_pure_python, generate_png
            if not os.path.exists(icon_ico):
                generate_ico_pure_python(icon_ico)
            if not os.path.exists(icon_png):
                generate_png(icon_png, 256)
        except Exception as exc:
            print(f"Warning: could not generate icons: {exc}")

    pyinstaller_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--name",
        APP_NAME,
    ]

    # Per-OS icon handling
    if system == "Windows" and os.path.exists(icon_ico):
        pyinstaller_cmd.extend(["--icon", icon_ico])
        pyinstaller_cmd.extend(["--add-data", f"{icon_ico}{os.pathsep}."])
    if os.path.exists(icon_png):
        pyinstaller_cmd.extend(["--add-data", f"{icon_png}{os.pathsep}."])

    pyinstaller_cmd.append("app.py")
    subprocess.run(pyinstaller_cmd, check=True)

    binary_name = APP_NAME + (".exe" if system == "Windows" else "")
    binary_path = os.path.join("dist", binary_name)
    if not os.path.exists(binary_path):
        raise FileNotFoundError(f"Expected binary not found: {binary_path}")

    zip_name = f"{APP_NAME}-{OS_LABELS[system]}.zip"
    zip_path = os.path.join("dist", zip_name)
    remove_path(zip_path)

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(binary_path, arcname=binary_name)

    print("Build complete.")
    print(f"Binary: {binary_path}")
    print(f"Zip: {zip_path}")

    # On Windows, also create a direct-download portable .exe with a clean name
    if system == "Windows":
        direct_exe = os.path.join("dist",
                                  f"{APP_NAME}-windows-portable.exe")
        shutil.copy2(binary_path, direct_exe)
        print(f"Direct download: {direct_exe}")

    # Build Windows installer with Inno Setup if available (optional)
    if system == "Windows":
        iss_path = os.path.join(root_dir, "installer.iss")
        if os.path.exists(iss_path):
            iscc = _find_inno_setup()
            if iscc:
                print(f"Building installer with: {iscc}")
                try:
                    subprocess.run([iscc, iss_path], check=True)
                    setup_path = os.path.join("dist", f"{APP_NAME}-Setup.exe")
                    if os.path.exists(setup_path):
                        print(f"Installer: {setup_path}")
                except Exception as exc:
                    print(f"Inno Setup failed (optional): {exc}")
            else:
                print("Inno Setup not found - skipping installer (optional)")


def _find_inno_setup():
    """Find Inno Setup compiler (iscc.exe) on the system."""
    # Check PATH first
    for name in ("iscc", "iscc.exe", "ISCC.exe"):
        path = shutil.which(name)
        if path:
            return path
    # Check common install locations
    for prog in (os.environ.get("ProgramFiles(x86)", ""),
                 os.environ.get("ProgramFiles", ""),
                 r"C:\Program Files (x86)",
                 r"C:\Program Files"):
        candidate = os.path.join(prog, "Inno Setup 6", "ISCC.exe")
        if os.path.exists(candidate):
            return candidate
    return None


if __name__ == "__main__":
    main()
