import importlib.util
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox


REQUIRED_DEPENDENCIES = {"customtkinter": "customtkinter", "packaging": "packaging", "openpyxl": "openpyxl", "beautifulsoup4": "bs4"}


def check_required_dependencies() -> None:
    missing = [package for package, module in REQUIRED_DEPENDENCIES.items() if importlib.util.find_spec(module) is None]
    if not missing:
        return
    requirements = Path(__file__).with_name("requirements.txt")
    install_command = f'"{sys.executable}" -m pip install -r "{requirements}"' if requirements.exists() else f'"{sys.executable}" -m pip install ' + " ".join(missing)
    message = f"Missing packages: {', '.join(missing)}\n\nInstall with:\n{install_command}"
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Missing Dependencies", message)
        root.destroy()
    except tk.TclError:
        print(message, file=sys.stderr)
    raise SystemExit(1)


def main() -> None:
    check_required_dependencies()
    from ui import run_app
    run_app()


if __name__ == "__main__":
    main()
