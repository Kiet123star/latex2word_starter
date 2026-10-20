"""Entry point: python run_server.py

Starts the FastAPI backend and serves the frontend at http://localhost:8000
"""
import os
import sys
import uvicorn

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    # Refresh PATH from Windows Registry to ensure recently installed tools (Pandoc, MiKTeX, Poppler) are in PATH
    try:
        import winreg
        paths = []
        for hkey in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            sub = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment" if hkey == winreg.HKEY_LOCAL_MACHINE else r"Environment"
            try:
                with winreg.OpenKey(hkey, sub) as key:
                    p, _ = winreg.QueryValueEx(key, "Path")
                    paths.append(p)
            except Exception:
                pass
        if paths:
            combined = ";".join(paths) + ";" + os.environ.get("PATH", "")
            # Deduplicate while preserving order
            seen = set()
            clean_paths = []
            for part in combined.split(";"):
                part_clean = part.strip()
                if part_clean and part_clean.lower() not in seen:
                    seen.add(part_clean.lower())
                    clean_paths.append(part_clean)
            os.environ["PATH"] = ";".join(clean_paths)
    except Exception:
        pass

if __name__ == "__main__":
    print("=" * 60)
    print("  LaTeX -> Word Converter")
    print("  Backend API:  http://localhost:8000/api/status")
    print("  Frontend UI:  http://localhost:8000")
    print("=" * 60)
    uvicorn.run(
        "backend.server:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info",
    )
