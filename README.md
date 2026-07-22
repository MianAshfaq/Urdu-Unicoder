# Urdu Unicoder

<p align="center">
  <img src="assets/urdu-unicoder-logo-final.png" alt="Urdu Unicoder logo" width="220">
</p>

Urdu Unicoder is an open-source Windows desktop application for recovering editable Unicode Urdu from legacy, text-based PDFs, rebuilding natural paragraphs, editing the result, designing a book layout, and exporting print PDF or UTF-8 HTML.

Created and maintained by **Muhammad Ashfaq**.

## Author and social links

- Website: [CyberOly.com](https://cyberoly.com/)
- GitHub: [@MianAshfaq](https://github.com/MianAshfaq)
- Facebook: [@MianAshfaq012](https://www.facebook.com/MianAshfaq012)

## Features

- Extracts selectable page ranges without freezing the interface
- Previews original PDF pages and compares source text with reconstructed text
- Normalizes Arabic Presentation Forms into searchable Unicode Urdu
- Recovers legacy PDFs that store Urdu words in reversed visual order
- Removes repeated `جاری ہے` page-continuation markers
- Rebuilds wrapped PDF rows into flowing paragraphs
- Provides a full RTL editor and live book preview
- Supports A5, B5, and A4 paper; inner/outer/top/bottom margins; running headers; and page numbers
- Controls Nastaleeq font, size, line height, paragraph spacing, indentation, word spacing, letter spacing, and alignment
- Includes manual and automatic paragraph cleanup tools
- Saves and reopens `.ubp` projects
- Exports print PDF and standards-based UTF-8 RTL HTML
- Includes tooltips, keyboard shortcuts, an F1 user guide, and an About/credits screen
- Includes an advanced Unicode editor toolbar with undo/redo, clipboard actions, search and replace, go-to-line, normalization, zoom, direction controls, and live statistics
- Imports Microsoft Word `.docx`, UTF-8/UTF-16 text, and Markdown files into the editor
- Pastes clipboard text while converting legacy presentation glyphs to standard Unicode Urdu
- Saves the editor as portable UTF-8 Unicode text
- Uses a dedicated Windows AppUserModelID and branded ICO so the running taskbar window shows Urdu Unicoder instead of Python
- Checks GitHub automatically for new versions and provides **Help → Check for Updates**

## Install on Windows

Requirements: Windows 10/11, Python 3.11–3.13, and an internet connection during initial setup.

1. Download or clone this repository.
2. Double-click `setup_windows.bat` to create a local virtual environment and install dependencies.
3. Double-click `run_windows.bat` to launch Urdu Unicoder.
4. Optional: run `build_exe_windows.bat` to create a standalone executable in `dist\Urdu Unicoder`.

For traditional Urdu book output, install **Noto Nastaliq Urdu**, **Jameel Noori Nastaleeq**, or **Mehr Nastaliq Web** in Windows.

## Recommended workflow

1. Choose **Open PDF** (`Ctrl+O`) and select an editable/text-based PDF.
2. Set **Start** and **End**. For a large book, test 10–30 pages first.
3. Keep both Unicode recovery options enabled for a legacy PDF, then choose **Extract and Reconstruct** (`Ctrl+R`).
4. Compare source and reconstructed text in **Text Editor** and make corrections.
5. Use the paragraph tools if the source contains hard line wraps.
6. Adjust layout settings and inspect **Book Preview**.
7. Save the editable project (`Ctrl+S`).
8. Export a sample PDF and inspect its Urdu shaping, margins, and page flow before exporting the whole book.

## Option reference

### Source and reconstruction

| Option | What it does | When to use it |
|---|---|---|
| Start / End | Limits extraction to an inclusive PDF page range. | Use a small range for testing or a chapter at a time. |
| Convert legacy Urdu glyphs to Unicode | Normalizes Arabic Presentation Form characters into standard Unicode. | Keep enabled for old Urdu publishing PDFs. |
| Recover reversed visual-order lines | Changes visually stored word order into logical Urdu reading order. | Enable when extracted words appear reversed; disable if correct text becomes reversed. |
| Remove repeated `جاری ہے` | Deletes page-end continuation markers. | Enable for novels/books that repeat this marker on pages. |
| Join wrapped lines | Treats PDF visual rows as parts of the same paragraph. | Usually enable for prose and novels. |
| Preserve blank lines | Keeps meaningful empty source lines. | Useful when continuous paragraph mode is disabled. |
| Join all wrapped lines into paragraphs | Ignores artificial blank rows inserted by some PDF generators. | Recommended for legacy novels with one extracted row per printed line. |

Scanned/image-only PDFs contain no editable text. Run OCR in another tool before opening them in Urdu Unicoder.

### Book layout

| Option | What it does | Practical guidance |
|---|---|---|
| Urdu font | Chooses an installed Unicode typeface. | Noto Nastaliq Urdu is a strong default. |
| Font size | Sets body text size in points. | Start around 14–18 pt for Nastaleeq, then print a sample. |
| Line height | Sets the line-height multiplier. | Nastaleeq often needs 1.8–2.3 to avoid collisions. |
| Paragraph spacing | Adds space after paragraphs. | Use a small value for novels; more for articles. |
| Page size | Selects A5, B5, or A4. | A5/B5 are common book sizes; A4 is useful for drafts. |
| Running header | Prints optional text at the page top. | Enter a book or chapter title, or leave blank. |
| First-line indent | Indents the first line in millimetres. | Use instead of large paragraph gaps for traditional books. |
| Word spacing | Fine-tunes space between words. | Keep near zero and adjust only after checking preview/PDF. |
| Letter spacing | Fine-tunes character spacing. | Keep at zero for Urdu unless a tested font needs adjustment. |
| Text alignment | Justifies, right-aligns, or centers text. | Justify for prose; center headings or special text. |
| Page numbers | Adds automatic PDF page numbers. | Disable for unnumbered front matter or special exports. |

### Margins

- **Top / Bottom:** vertical whitespace and room for header/footer content.
- **Inner:** binding-side margin; normally wider so text is not lost in the spine.
- **Outer:** outside edge margin.

### Paragraph tools

Tools process the selected editor text. If nothing is selected, they process the complete document.

- **Join Selected Lines as Paragraph:** removes hard wraps and keeps blank-line paragraph boundaries.
- **Split Selected Text at ۔ or .:** inserts paragraph breaks after sentence-ending punctuation.
- **Minimum source lines:** controls how many source rows automatic paragraph creation waits before accepting a sentence ending.
- **Auto Make Paragraphs:** joins text and begins a new paragraph only after the minimum line count and sentence punctuation are reached.
- **Remove Extra Line Breaks:** removes line wrapping inside paragraphs while preserving actual blank-line paragraph breaks.

## Keyboard shortcuts

| Shortcut | Command |
|---|---|
| `Ctrl+N` | New project |
| `Ctrl+O` | Open PDF |
| `Ctrl+Shift+O` | Open project |
| `Ctrl+S` | Save project |
| `Ctrl+R` | Extract and reconstruct |
| `Ctrl+Shift+P` | Export PDF |
| `F1` | Complete user guide |
| `Ctrl+F` | Find and replace |
| `Ctrl+G` | Go to line |
| `Ctrl+Shift+U` | Normalize selected text or the complete document |
| `Ctrl+Shift+V` | Paste clipboard text and convert it to Unicode |
| `Ctrl+Shift+I` | Text statistics |
| `Ctrl+Alt+W` | Import a Microsoft Word `.docx` file |
| `Ctrl+Alt+O` | Import a text or Markdown file |
| `Ctrl+Alt+S` | Save editor content as UTF-8 text |
| `Ctrl+D` | Duplicate current line |
| `Ctrl+Shift+K` | Delete current line |
| `Ctrl++` / `Ctrl+-` | Editor zoom in/out |
| `Ctrl+0` | Reset editor zoom |

## Advanced text editor

The Text Editor tab provides a Unicode-safe editing toolbar while keeping the document as portable plain text:

- **Undo / Redo, Cut / Copy / Paste, Select All:** familiar document-editing commands with standard Windows shortcuts.
- **Find / Replace:** forward and backward search, case-sensitive matching, whole-word matching, single replacement, and replace-all.
- **Go to Line:** jumps directly to a line in very large books.
- **Normalize Unicode:** converts selected text—or the complete document when nothing is selected—to normalized Unicode compatibility form.
- **Text Statistics:** reports words, characters, characters without spaces, lines, and paragraphs.
- **Zoom:** increases or decreases only the editing view; it does not change the exported book font size.
- **Right-to-Left / Left-to-Right:** switches editor direction for Urdu or mixed-language material.
- **Live status:** continuously shows the cursor line/column and document word/character counts.
- **Paste + Unicode:** reads text from the Windows clipboard, normalizes legacy Urdu presentation glyphs, removes directional control artifacts, and inserts editable Unicode text.
- **Import Word Document:** imports `.docx` paragraphs, headings, lists, and table rows in document order. Word styling is converted to editable plain text for reliable Urdu publishing.
- **Import Text File / Save Editor Text:** reads common Unicode encodings and saves portable UTF-8 text.
- **Recover Legacy Visual Order:** repairs selected or complete legacy text whose Urdu words are stored in display order.
- **Clean Whitespace:** removes repeated spaces, trailing spaces, and excessive blank lines without merging paragraphs.
- **Duplicate / Delete Line:** provides fast manuscript line editing with `Ctrl+D` and `Ctrl+Shift+K`.

## Limitations

- OCR is not included.
- Reconstruction is rule-based; ambiguous material requires editorial review.
- Output quality depends on the Unicode Urdu font installed on the computer.
- PDF inner/outer margins are applied to the exported layout but are not alternated automatically on odd/even pages.
- DOCX and EPUB export are not currently included.

## Development and tests

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m unittest discover -v
.venv\Scripts\python.exe app\main.py
```

## Updates

Urdu Unicoder checks the public `version.json` manifest on this repository shortly after startup. It does not install anything silently.

- When a newer version is available, the user sees the version number and release notes and chooses whether to install it.
- Source installations downloaded or cloned from GitHub use the external updater, which closes the application, downloads the official repository archive over HTTPS, updates only application files, installs new Python requirements, and restarts Urdu Unicoder.
- `.ubp` projects, exported books, the local virtual environment, and Git history are not replaced by the updater.
- Packaged EXE builds open the official GitHub Releases page for the new installer because a running Windows executable should not replace itself.
- Use **Help → Check for Updates** to run a manual check at any time.

Maintainers must update `APP_VERSION` in `app/main.py` and the matching `version` and release notes in `version.json` in the same commit. The automated tests fail if these versions differ.

## License

Urdu Unicoder is released under the [MIT License](LICENSE.txt). Third-party packages keep their respective licenses.
