# excel_backend.py
import os
import platform

def excel_backend() -> str:
    """
    Returns:
        'windows' -> Excel COM + PowerShell allowed
        'cloud'   -> Streamlit Cloud / Linux / macOS (no Excel COM)
    """
    if platform.system() == "Windows" and not os.environ.get("STREAMLIT_RUNTIME"):
        return "windows"
    return "cloud"