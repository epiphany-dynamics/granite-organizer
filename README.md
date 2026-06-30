# Granite Net Organizer

A desktop tool (Python + Tkinter) built for CIPP/sewer rehabilitation inspection workflows. It sorts inspection PDFs and videos into the correct segment folders by matching manhole IDs found in filenames against an Excel prep sheet, replacing manual file sorting on multi-hundred-file inspection jobs.

Built by [Patrick Gibbs](https://github.com/patrickg21212) / [Epiphany Dynamics](https://epiphanydynamics.ai), drawing on direct field experience in sewer inspection.

## What it does

- Reads a mainline/lateral lookup table from an Excel prep sheet (matches segment IDs and manhole IDs across all sheets)
- Scans a folder of inspection PDFs and videos
- Matches each file to its segment by manhole ID pattern (e.g. `123-45-678`)
- Moves each file into the correct segment folder automatically

## Running it

Requires Python 3.12 and `openpyxl`:

```sh
pip install openpyxl
python granite_organizer.py
```

## Windows build

Every push to `main` builds a standalone `GraniteNetOrganizer.exe` via GitHub Actions (PyInstaller, `--onefile --windowed`), available as a workflow artifact, no Python install needed on the end-user's machine.
