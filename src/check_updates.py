import requests
import sys
import subprocess
import tempfile
from pathlib import Path

from config import VERSION, REPO


def check_for_updates() -> bool:
    """
    Checks GitHub Releases for a newer version.
    Returns True if an update was started.
    """
    try:
        response = requests.get(
            f"https://api.github.com/repos/{REPO}/releases/latest",
            timeout=15
        )
        response.raise_for_status()

        data = response.json()
        latest_version = data["tag_name"].lstrip("v")

        if is_newer_version(latest_version, VERSION):
            print(f"New version {latest_version} available (current: {VERSION})")

            assets = data.get("assets", [])
            if not assets:
                raise RuntimeError("No downloadable assets found in release")

            download_url = assets[0]["browser_download_url"]
            update_app(download_url)
            return True

    except Exception as e:
        print(f"Update check failed: {e}")

    return False


def update_app(download_url: str):
    """
    Safely downloads and installs an updated executable on Windows.
    """

    exe_path = Path(sys.executable)
    exe_name = exe_path.name

    temp_dir = Path(tempfile.gettempdir())
    new_exe = temp_dir / f"{exe_name}.new"
    updater_bat = temp_dir / "centrofit_updater.bat"

    try:
        # Download new executable (streamed)
        response = requests.get(download_url, stream=True, timeout=60)
        response.raise_for_status()

        with open(new_exe, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        if new_exe.stat().st_size < 1_000_000:
            raise RuntimeError("Downloaded file is too small to be a valid executable")

        # Create updater batch file
        updater_bat.write_text(
            f"""@echo off
set EXE_PATH="{exe_path}"
set NEW_EXE="{new_exe}"

timeout /t 3 /nobreak > nul

if not exist %NEW_EXE% (
    exit /b 1
)

del /f /q %EXE_PATH%
move /y %NEW_EXE% %EXE_PATH%

start "" %EXE_PATH%
del "%~f0"
""",
            encoding="utf-8"
        )

        # Run updater and exit app
        subprocess.Popen(
            ["cmd.exe", "/c", str(updater_bat)],
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )

        sys.exit(0)

    except Exception as e:
        print(f"Update failed: {e}")
        if new_exe.exists():
            new_exe.unlink(missing_ok=True)


def is_newer_version(latest: str, current: str) -> bool:
    """
    Compares semantic versions correctly.
    Example: 1.10.0 > 1.2.0
    """
    def parse(v):
        return tuple(int(x) for x in v.split("."))

    return parse(latest) > parse(current)
