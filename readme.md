# Post Utilities

Small Python tools for assistant editing on documentary features and TV series: renaming camera originals to match an edit, filling in media details in a tracking spreadsheet, and turning EDLs into music cue lists.

They were written for real projects on macOS and are shared as-is. Expect practical scripts rather than polished software.

## What's here

| Tool | What it does | Status |
|---|---|---|
| [Rename Files from Excel](#rename-files-from-excel) | Renames media files using a spreadsheet of old and new names. Command line and GUI. | Working, has tests |
| [MX EDL Parser](#mx-edl-parser) | Turns a CMX 3600 EDL into a clip list (e.g. a music cue sheet) as text, CSV or Excel. | Working |
| [Add Media Stats to Excel](#add-media-stats-to-excel) | Fills resolution, codec, frame rate, duration, start timecode and size into a spreadsheet. | Rough script |
| [Generate Test Files](#generate-test-files) | Creates empty placeholder files from a spreadsheet, for trying out a rename. | Rough script |
| [Comprehensive Media Detector](#comprehensive-media-detector-concept-only) | Black frame, flash frame and silence detection. | **Concept only, doesn't work** |

Several tools that used to live here now have their own repositories. See [Tools that moved](#tools-that-moved).

## Requirements

- **Python 3.10 or newer**
- **Python packages**, depending on the tool (listed in each section below)
- **FFmpeg** (`ffprobe`) for *Add Media Stats to Excel*. On macOS: `brew install ffmpeg`

To install everything at once:

```bash
pip install pandas openpyxl PyQt5 timecode pycmx
```

---

## Rename Files from Excel

**Folder:** `rename_files_from_excel/`

Renames files to match names from a spreadsheet. A typical use is renaming camera originals to the names their proxies were given in Avid, so the two can be matched up.

### How it works

The spreadsheet needs two columns:

| Master Original Name | Avid Proxy Name |
|---|---|
| A001C003_250101_R1AB | Interview_Jane_01 |
| A001C004_250101_R1AB | Interview_Jane_02 |

For each file in the target folder whose name (without its extension) matches a cell in **Master Original Name**, the file is renamed to the **Avid Proxy Name**, plus a suffix, plus its original extension. With the default suffix `_M`:

```
A001C003_250101_R1AB.mov  →  Interview_Jane_01_M.mov
```

- The spreadsheet must have a single sheet.
- Stray spaces in cells are ignored.
- Before anything is renamed, the whole spreadsheet is checked: empty cells, duplicate names, non-ASCII characters, and characters that aren't allowed in file names stop the run with the row number.
- Subfolders are searched by default. Hidden files and folders (such as `.Trashes`) are skipped.
- If a file with the new name already exists, it is skipped, unless you use `--force`. With `--force`, the existing file is renamed to a timestamped backup first (e.g. `Interview_Jane_01_M_backup_20251027140501.mov`).
- Every rename is written to a log file, `renames.log`.

### Command line

```bash
pip install pandas openpyxl

python rename_files_from_excel/rename_files_from_excel.py names.xlsx /Volumes/Media/Day01 --dry-run
```

Always do a `--dry-run` first: it lists every rename it would make without changing anything.

| Option | Default | Meaning |
|---|---|---|
| `--files-col` | `Master Original Name` | Column with the current file names |
| `--rename-col` | `Avid Proxy Name` | Column with the new names |
| `--suffix` | `_M` | Added to the end of every new name (use `--suffix ""` for none) |
| `--no-recursive` | off | Only look in the top folder, not subfolders |
| `--force` | off | Replace existing files, keeping the old one as a backup |
| `--dry-run` | off | Show what would happen without renaming anything |
| `--log` | `renames.log` | Log file location (relative to the folder you run from) |

### GUI

```bash
pip install pandas openpyxl PyQt5

python rename_files_from_excel/rename_files_from_excel_gui.py
```

The GUI has the same options, and can rename across several folders in one go. Its log is saved next to the script.

### Tests

```bash
pip install pytest
cd rename_files_from_excel
python -m pytest
```

---

## MX EDL Parser

**Folder:** `MX_edl_parser/`

Reads a CMX 3600 EDL and lists every clip used, with source and record timecodes. It was built for music cue sheets, but works for any EDL.

- Uses of the same clip separated by a gap of more than one frame are listed as separate instances.
- Writes text, CSV or Excel files.
- `--analytics` adds a report on how much of the timeline each clip takes up.

```bash
pip install -r MX_edl_parser/requirements.txt

python MX_edl_parser/MX_edl_parser.py music.edl --format excel --fps 23.976 --analytics
```

| Option | Default | Meaning |
|---|---|---|
| `--format` | `txt` | `txt`, `csv`, `excel` or `all` |
| `--fps` | `23.976` | Frame rate for timecode maths |
| `--analytics` | off | Also write a usage report |
| `--no-overwrite` | off | Don't replace existing output files |

`pycmx` is used for parsing when installed. Without it, the parser falls back to a built-in one, and the text and CSV formats need no packages at all. Excel output needs `pandas` and `openpyxl`. See the [EDL parser README](MX_edl_parser/README.md) for details.

---

## Add Media Stats to Excel

**File:** `add_media_stats_to_excel.py`

Fills in technical details for each file listed in a tracking spreadsheet, using `ffprobe`. It finds files whose names (without extension) match the **Master Original Name** column, searching subfolders, and writes these columns:

`Master File Size (GB)`, `Master Resolution`, `Master Codec`, `Master FPS`, `Master Duration`, `Master Start`

```bash
pip install pandas openpyxl timecode

python add_media_stats_to_excel.py tracking.xlsx /Volumes/Media updated_tracking.xlsx
```

The arguments are, in order: the input spreadsheet, the media folder, and the output spreadsheet (which must end in `.xlsx`).

This is a rough script. It takes no options, has no `--help`, and the column names are fixed at the top of the file. It is the companion to *Rename Files from Excel* and expects the same spreadsheet layout.

---

## Generate Test Files

**File:** `generate_test_files.py`

Creates an empty `.mov` file for every name in the **Master Original Name** column. This gives you a safe set of dummy files to try *Rename Files from Excel* on before touching real media.

```bash
pip install pandas openpyxl

python generate_test_files.py names.xlsx /path/to/empty/test_folder
```

The target folder must already exist.

---

## Comprehensive Media Detector (concept only)

**Folder:** `comprehensive_media_detector/`

> [!WARNING]
> **This tool doesn't work.** It's an unfinished concept and shouldn't be relied on for QC.

The idea is to use FFmpeg to find black frames, flash frames and audio dropouts in a video file, and report them with timecodes as text, CSV or Excel. The code is here as a starting point.

`user-guide.md` describes how the tool is *meant* to work. It is a design document, not instructions for something finished.

---

## Tools that moved

These used to be part of this repository and now have their own:

| Tool | Repository |
|---|---|
| Directory Indexer: browsable HTML archives of drive contents | [snorkem/directory-indexer](https://github.com/snorkem/directory-indexer) |
| Copy File List: find and copy files from a list or EDL | [snorkem/copy-file-list](https://github.com/snorkem/copy-file-list) |
| Lower Thirds from Excel: lower-third graphics from a spreadsheet | [snorkem/l3rds-from-excel](https://github.com/snorkem/l3rds-from-excel) |
| ALE Converter, Sony XML to Avid Markers, MediaSilo to Avid Markers | [snorkem/avid-media-composer-scripts](https://github.com/snorkem/avid-media-composer-scripts) |

## Notes

- Developed and used on macOS. The scripts are plain Python and should run elsewhere, but Windows and Linux are much less tested.
- Try a dry run, or work on copies, before pointing anything at real media. Renames can't be undone automatically.
