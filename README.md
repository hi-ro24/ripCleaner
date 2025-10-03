ripCleaner version 0.5

Summary:
- Monitors RIP output folders and removes temporary TIFF files matching the pattern:
  bip<0-5>-output-1bpp-<page_number>.tif
- Designed to avoid deleting files that are being written and to minimize network load.

Usage:
- Configure paths and options in config.ini placed next to the executable.
- Run:
    ripCleaner.exe --version
    ripCleaner.exe --kick RIP1
    ripCleaner.exe         # polling mode

Notes:
- Logging is required. If the log directory cannot be created or written, the program exits with an error.
- For Windows, QuickEdit mode is disabled at startup to prevent accidental pause by console selection.
