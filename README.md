# BlockSearch (DocSearch)
FOR USE WITH VERBATIM
A small **Python + Tkinter** desktop app that indexes **Microsoft Word `.docx`** files by heading hierarchy, lets you **search** across sections, and on **macOS** can drive **Word** to copy ranges, jump to a heading, or run a VBA macro.

## What it does

- Loads one or more `.docx` files and parses structure from Word’s built-in heading styles.
- Maps styles to four levels (labels in the UI: **pocket**, **hat**, **block**, **tag**):

  | Word style   | Level label |
  | ------------ | ----------- |
  | `Heading1`   | pocket      |
  | `Heading2`   | hat         |
  | `Heading3`   | block       |
  | `Heading4`   | tag         |

- Shows a searchable, virtualized list of sections with previews; filter by level and by source document.
- **macOS + Word**: optional actions use **AppleScript** against Microsoft Word (open at heading, copy a section to the clipboard, run the macro named `SendToSpeechCursor` from the UI).

## Requirements

- **Python 3** with the standard library (uses `tkinter`, `zipfile`, `xml.etree`, `threading`, etc.).
- **macOS** for Word automation (AppleScript). Search and browsing inside the app do not require Word; copy / open / send actions do.
- **Microsoft Word** installed when using open, copy, or VBA send features.

## Run

From this directory:

```bash
python3 docsearch.py
```

## Optional auto-load

On startup, the app can look under `~/debate` and load `.docx` files from configured subfolders (e.g. topic-specific vs general). Paths are set near the top of `docsearch.py` (`AUTO_LOAD_FOLDER`, `DEBATE_TOPIC_SPECIFIC_SUBDIRS`, `DEBATE_GENERAL_SUBDIRS`). Set `AUTO_LOAD_FOLDER` to `None` to disable.

## VBA macro name

The “send” action types the macro name configured as `SEND_VBA_MACRO_NAME` in `docsearch.py` (default: `SendToSpeechCursor`). That macro must exist in your Word environment.

## `save/` folder

Contains earlier iterations of the script (`v1.py` … `v7.py`) kept for reference; the current application is **`docsearch.py`**.

## License

No license is specified in this repository; add one if you intend to distribute or accept contributions.
