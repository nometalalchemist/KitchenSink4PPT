"""COM layer: operations that need a running PowerPoint (Windows only).

`bridge` holds the singleton-safe application context plus the export and
validation operations built on it. Heavy imports (pythoncom, win32com,
winreg) happen inside function bodies, never at module top, so the
file-based server imports this package safely on any platform. The future
live layer (live.py, live_ops.py) bolts on beside it.
"""
