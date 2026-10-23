# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Python CLI that styles chapter titles as Heading 1 (centered, black) in `.doc`/`.docx`/`.odt` files and always writes `.odt`. The user writes in Hungarian; code comments, messages and docs are Hungarian too.

## Commands

```bash
python3 -m venv .venv && .venv/bin/pip install -e .   # setup (python-docx is the only dependency)
.venv/bin/fejezet-formazo <folder>              # output: <folder>/formazott/*.odt, originals deleted
.venv/bin/fejezet-formazo <folder> --megtart    # keep originals
```

There are no automated tests. Verify changes on copies in a scratch folder **inside the home directory, not under `/tmp`**: LibreOffice is the Flatpak build (`org.libreoffice.LibreOffice`), which cannot see the host `/tmp`. The system-wide `soffice` lacks Writer and cannot convert. Never run the tool on the user's own book folders without `--megtart` unless asked: originals are deleted after conversion. The default input folder is `../test`.

The command is also installed user-wide in `~/.local/bin/fejezet-formazo` (editable install with the system Python), so edits to `main.py` apply to both.

## Architecture

Everything is in `fejezet_formazo/main.py`. The pipeline for each file:

1. If the input is `.doc`/`.odt`, LibreOffice converts it to `.docx`.
2. python-docx finds chapter paragraphs and applies Heading 1 (set both on the style and directly on the paragraph and its runs, so direct formatting can't override it).
3. LibreOffice converts the result to `.odt` in the output folder.

Temporary files and LibreOffice profiles live inside the output folder for the Flatpak reason above.

Behaviour that is easy to break:
- **Chapter detection** has two parts. `CHAPTER_PATTERN` matches "N. fejezet", "<word> fejezet" and the `FRONT_BACK_MATTER` names (Előszó, Epilógus, Utóhang, ...). Bare numbers ("1.", "II.") count only as part of the longest strictly consecutive 1, 2, 3, ... run (`longest_numbered_sequence`, minimum 2 items); out-of-order numbers are skipped. Roman numerals require the trailing dot, because drop-cap paragraphs contain a lone "I".
- **Existing headings:** paragraphs that already have another heading style (e.g. a Heading 2 "Utóhang") keep their level. A paragraph's style can be `None` (it references a style missing from the file).
- **LibreOffice hangs:** LibreOffice loops forever on some files. Every conversion has a `CONVERT_TIMEOUT` and its own temporary profile (`-env:UserInstallation`). On timeout, processes are killed by matching that profile path in `/proc/*/cmdline`, because the sandboxed `soffice.bin` outlives the killed `flatpak run`. Never use `flatpak kill`, which would close the user's open LibreOffice.
- **Deleting originals:** the original is deleted only after the output passes `is_valid_odt`. A second input with the same stem in one run (`x.doc` + `x.odt`) is rejected so its output cannot overwrite the first. One failing file must never stop the batch.
