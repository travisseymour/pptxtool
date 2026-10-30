# PPTX Tool

Travis L. Seymour, PhD. 2006

A GUI application for turning PowerPoint decks into printer-friendly notes and watermarked, compressed PDF handouts.

[![PPTX Tool UI](pptxtool/resources/pptxtool_ui.png)](pptxtool/resources/pptxtool_ui.png)

---

## Features

Choose a file with **Browse…** or drag it onto the window. The available functions depend on the file:

| File              | Functions                                               |
| ----------------- | ------------------------------------------------------- |
| `deck.pptx`       | PDF, PDF + Watermark + Compress, Pre-Notes, Standardize |
| `deck_notes.pptx` | PDF, Post-Notes, Standardize                            |
| `file.pdf`        | Compress, Watermark                                     |

- **PDF**: convert to PDF with LibreOffice.
- **Pre-Notes**: write a printer-friendly `<stem>_notes.pptx` (black Arial text, black round bullets, white backgrounds, no master decorations or hidden slides).
- **Post-Notes** / **PDF + Watermark + Compress**: PDF, compress, then write a watermarked `<stem>_reader.pdf` (selectable text) and a raster `<stem>.pdf`.
- **Standardize**: Arial fonts and black round bullets, edited in place.
- **Compress**: shrink a PDF with Ghostscript.
- **Watermark**: write reader and raster watermarked copies of a PDF.

Each function opens a dialog that describes what will happen, shows its settings and output files, and asks before overwriting anything. Progress appears in the Status area. The app follows the system light/dark setting.

You can also pass a file on the command line: `pptxtool deck.pptx`.

## System Dependencies

- [LibreOffice](https://www.libreoffice.org/) (`libreoffice` or `soffice` on PATH) for PDF conversion
- [Ghostscript](https://ghostscript.com/) (`gs` on PATH) for compression

On Debian/Ubuntu: `sudo apt install libreoffice-impress ghostscript`

## Installation

#### Install uv

- Installing uv on linux and macos

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

- Installing uv on Windows

```bash
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

#### Install PPTX Tool

```bash
uv tool install git+https://github.com/travisseymour/pptxtool
# or, from a local checkout:
uv tool install .
```

### Update PPTX Tool

```bash
uv tool upgrade pptxtool
```

### Uninstall PPTX Tool

```bash
uv tool uninstall pptxtool
```

## License

GPLv3+
