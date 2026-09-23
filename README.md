# pub2pptx

A single-file command line tool that converts Microsoft Publisher (`.pub`) files
to PowerPoint (`.pptx`).

Each Publisher page becomes one slide of the same size. In the default
*editable* mode, pictures, shapes, text boxes and tables are rebuilt as native
PowerPoint objects with their fonts, sizes, colours, alignment and positions, so
the result can be edited in PowerPoint. A *snapshot* mode paints each page into
a single picture per slide instead.

## Requirements

- [uv](https://docs.astral.sh/uv/) (the script declares its Python dependencies
  inline, PEP 723, so no manual `pip install` is needed; it needs Python 3.12
  or newer, which uv fetches automatically if none is installed)
- `pub2raw` from **libmspub**, which reads the Publisher file
  - Arch: `sudo pacman -S libmspub`
  - Debian/Ubuntu: `sudo apt install libmspub-tools`
  - macOS: `brew install libmspub`
  - Windows: see below

Nothing else is needed. Fontconfig's `fc-match`, if present, lets the
thumbnail and snapshot painter use the document's real fonts; without it the
painter looks the fonts up in the system font folders itself.

### Windows

The script runs unchanged on Windows with `uv run pub2pptx.py file.pub`.
`pub2raw.exe` comes from [MSYS2](https://www.msys2.org/): in an MSYS2 shell run

```sh
pacman -S mingw-w64-ucrt-x86_64-libmspub
```

The tool looks for it under `C:\msys64` (ucrt64, mingw64 or clang64) as well
as on the `PATH`; for any other location set the `PUB2RAW` environment
variable to the full path of `pub2raw.exe`.

## Usage

```sh
uv run pub2pptx.py brochure.pub                 # writes brochure.pptx next to it
uv run pub2pptx.py brochure.pub -o out/deck.pptx
uv run pub2pptx.py *.pub -o converted/          # several files into a folder
uv run pub2pptx.py menu.pub --font-map "Raleway=Calibri,Poppins=Segoe UI"
uv run pub2pptx.py menu.pub --mode snapshot --dpi 200
```

The script is also executable directly (`./pub2pptx.py file.pub`) thanks to its
`uv run --script` shebang.

Options:

| Option | Meaning |
| --- | --- |
| `-o, --output` | Output file for a single input, or output directory for several |
| `--mode editable\|snapshot` | `editable` (default) builds native slides; `snapshot` paints one picture per page |
| `--dpi N` | Paint resolution in snapshot mode (default 150) |
| `--font-map A=B,...` | Substitute font names |
| `--blank-size PT` | Font size for blank spacer lines inside table cells (see below) |
| `--no-thumbnail` | Skip the embedded first-page preview (see below) |
| `--dump-raw` | Also write the `pub2raw` dump next to the output, for debugging |
| `-v` | Verbose logging |

## How it works

1. `pub2raw` (libmspub's test tool) prints the document as a stream of
   librevenge drawing callbacks: pages, styles, polygons with picture fills,
   paths, text objects with paragraphs and spans, and tables.
2. The script parses that stream into a small document model.
3. [python-pptx](https://python-pptx.readthedocs.io/) writes the slides.
   Rectangles with a stretched picture fill become Pictures (with rotation),
   tiled picture fills become picture-filled shapes, paths become freeform
   shapes (curves preserved, arcs flattened), text keeps font, size, bold,
   italic, underline, colour, small caps, super/subscript, alignment, indents,
   spacing and line breaks, and tables keep column widths, row heights, merged
   cells, fills and borders.

WMF and EMF clip art, which Pillow cannot decode, is embedded as-is with its
native content type; PowerPoint renders those formats itself. Such images are
left out of thumbnails and snapshots.

## Blank lines and row heights in tables

Publisher authors often position text inside a table cell by pressing Enter a
few times above it. The paragraph mark of such a blank line has its own font
size, and libmspub discards it for table cells (it keeps it for text boxes, so
those are exact). The script recovers a good estimate from the geometry:
Publisher keeps a row at its stored height unless the text overflows, so when
a table has not grown, the space left in a cell after its visible text, divided
by the number of blank lines, bounds their size. The tightest bound in the
table is used, capped at the size of the neighbouring text. Pass
`--blank-size` to override it when you know the real value.

Rows that did grow to fit their text are also handled: Publisher records the
growth only in the table's total height, so the difference between that and
the sum of the stored row heights is handed to the rows whose content
overflows, in proportion to how much they overflow.

## File-manager thumbnails

Each `.pptx` carries a 256px JPEG of its first page in the standard
`docProps/thumbnail.jpeg` slot, painted with Pillow straight from the converted
page (pictures, shapes, text and tables). This is the same mechanism PowerPoint
uses, and Linux file managers read it:

- **KDE / Dolphin**: built in via kio-extras. Enable it under
  *Configure Dolphin > Interface > Previews > Office files*.
- **GNOME / Nautilus**: needs libgsf's `gsf-office-thumbnailer`
  (Arch: `pacman -S libgsf`, Debian/Ubuntu: `apt install libgsf-bin`).

The same painter produces the slides in snapshot mode. It is a preview-grade
renderer: per-paragraph text styling, simple word wrapping, curves flattened,
gradients reduced to a single colour.

## Limitations

- Fonts are referenced by name; if a font used in the Publisher file is not
  installed, PowerPoint substitutes one (use `--font-map` to choose).
- Publisher features libmspub does not expose (text wrapping around shapes,
  linked text boxes that overflow, some effects and 3D/WordArt) are lost.
- A presentation has a single slide size, taken from the first page.
- Gradients are reduced to their first and last colour stops.
