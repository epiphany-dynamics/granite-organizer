# Granite Net Organizer

A desktop tool that sorts sewer and CIPP inspection PDFs and videos into the correct segment folders by matching manhole IDs in the filenames against an Excel prep sheet.

## Why this exists

Sorting several hundred inspection files into segment folders by hand takes hours and one misfiled inspection becomes a support call later. This tool reads the prep sheet once, places every file it can match into its segment folder in a single pass, and lists the ones it could not place along with the reason.

## Quickstart

Requires Python 3.12 and `openpyxl`:

```sh
pip install openpyxl
python granite_organizer.py
```

The window opens in a second. Pick the Excel prep sheet, the PDF folder, the output folder, and optionally a video folder, then press Organize Files. Each file is logged as it is placed. Leave "Move files instead of copy" unchecked to keep the originals untouched while you check the result.

Files land under `<output>/<segment>/`, or `<output>/<segment>/Lat-<tap>/` for lateral inspections.

## How it works

- Opens the prep sheet with `openpyxl` and scans every sheet, separating mainline rows (a segment ID plus two manhole IDs) from lateral rows (two manhole IDs plus a tap number, no segment ID).
- Matches each filename against the `123-45-678` manhole ID pattern, then looks that ID up in the mainline and lateral tables.
- Copies or moves the file into the segment folder for the match; a lateral also gets its own `Lat-<tap>` subfolder.
- Never overwrites: a name collision gets a `_1`, `_2` suffix.
- Logs every file as matched or unmatched, with the reason (no manhole IDs found, or a lateral tap whose parent segment was not in the sheet).
- Remembers the last folders used in `organizer_config.json`, saved next to the script.

## Tests

Tests: none yet. There is no test suite in this repo. The tool is checked by running it against a real prep sheet and reading the log.

## Known limits

- The manhole ID format is hardcoded to the 3-2-3 pattern (`123-45-678`). Other ID formats need a code change.
- Matching is filename based only. Nothing is read from inside the PDFs.
- `.github/workflows/build-exe.yml` is the PyInstaller recipe for a standalone `GraniteNetOrganizer.exe`, but GitHub Actions is disabled at the account level so it does not currently run. To build the executable by hand:

  ```sh
  pip install pyinstaller
  pyinstaller --onefile --windowed --name GraniteNetOrganizer granite_organizer.py
  ```

- No undo. The default is copy, so a wrong run leaves the originals where they were.

## License

MIT (see LICENSE).
