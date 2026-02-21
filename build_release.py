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
        "app.py",
    ]
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


if __name__ == "__main__":
    main()
