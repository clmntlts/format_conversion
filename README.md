# Questionnaire Long → Wide Converter

A desktop tool that converts optically-read questionnaire exports from
**long format** (one row per question / matrix item) to **wide format**
(one row per `submission_id`, one column per variable), ready for analysis
in Excel, SPSS, R, JASP, or Python.

The questionnaires are fully modular, so the tool does not assume a fixed set
of questions: it reads whatever question codes and matrix items are actually
present in each file and builds the variables from them.

---

## Quick start: download and use (Windows)

No installation, no Python, no administrator rights needed.

1. Open the **[latest release](https://github.com/clmntlts/format_conversion/releases/latest)** page.
2. Under **Assets**, click **`Questionnaire_Converter.exe`** to download it
   ([direct download link](https://github.com/clmntlts/format_conversion/releases/latest/download/Questionnaire_Converter.exe)).
   If your browser says the file "isn't commonly downloaded", choose **Keep**.
3. Move the file wherever you like (Desktop, a project folder…). It is a single,
   self-contained program.
4. **Double-click** it. The window takes **10–30 seconds** to appear (the
   program unpacks itself at each launch): wait, there is no need to
   double-click again. If Windows shows a blue **"Windows protected your PC"** dialog,
   click **More info → Run anyway** (the program is not code-signed).
5. In the window that opens:
   1. **Add files…** (or **Add folder…**) and pick your long-format `.xlsx` exports.
   2. Check the detected variables in the **2. Variables** tab and, if you want,
      double-click an **Export name** to rename a column.
   3. Click **Refresh preview** in the **3. Preview** tab to check the result.
   4. Click **Export wide .xlsx…** and choose where to save.

Each exported file has a `wide` sheet (the data, one row per respondent) and a
`codebook` sheet (what every column means). Details are in
[section 4](#4-step-by-step-usage) and [section 5](#5-the-output-file).

**Updating:** download the new `.exe` from the same page and replace the old one.

**macOS / Linux:** use the Python version, see [section 3](#3-running-the-tool).

---

## 1. What the tool does

### Long format (input)

The optical-reading tool produces one row **per answered item**. A single
respondent therefore spans dozens of rows — one for each simple question and
one for each row of every matrix question:

| submission_id | question_code | question_type | matrix_row | answer_1 |
|---------------|---------------|---------------|------------|----------|
| MwQ1A9kB-001  | 2db6OGV1sq    | coded_answer  | row_0      | AB       |
| MwQ1A9kB-001  | DZIkow9Bn5    | matrix        | a. …       | Presque jamais |
| MwQ1A9kB-001  | DZIkow9Bn5    | matrix        | b. …       | Souvent  |
| …             | …             | …             | …          | …        |

### Wide format (output)

The tool pivots this into one row **per submission**, with every item becoming
its own column:

| submission_id | 2db6OGV1sq | DZIkow9Bn5_a   | DZIkow9Bn5_b | … |
|---------------|-----------|----------------|--------------|---|
| MwQ1A9kB-001  | AB        | Presque jamais | Souvent      | … |
| MwQ1A9kB-002  | CD        | Jamais         | Toujours     | … |

This is the shape almost every statistics package expects: one case per line.

---

## 2. Requirements

- **Using the .exe:** Windows 10 or 11 (64-bit). Nothing else: it is
  self-contained and runs without Python installed.
- **Running from source (any OS):** Python 3.10+ with `pandas` and `openpyxl`.
  Install with:

  ```bash
  pip install -r requirements.txt
  ```

  `tkinter` (the GUI toolkit) ships with the standard Python installer on
  Windows and macOS; on Debian/Ubuntu install it with `sudo apt install python3-tk`.

---

## 3. Running the tool

- **Easiest — the executable (Windows):** download
  **`Questionnaire_Converter.exe`** from the
  [latest release](https://github.com/clmntlts/format_conversion/releases/latest)
  and double-click it (see [Quick start](#quick-start-download-and-use-windows)).
  The window takes 10–30 seconds to appear while the program unpacks itself.
- **From source:** get the code (`git clone https://github.com/clmntlts/format_conversion.git`,
  or **Code → Download ZIP** on GitHub and unzip it), then in that folder:

  ```bash
  pip install -r requirements.txt
  python Questionnaire_Converter.py
  ```

- **Scripted / batch (no GUI):** import `converter_core` — see section 8.

> **Windows SmartScreen note:** because the .exe is unsigned, Windows may show
> a blue “Windows protected your PC” dialog the first time. Click
> **More info → Run anyway**. This is expected for any self-built executable.

---

## 4. Step-by-step usage

1. **Add data** — *1. Data files* panel:
   - **Add files…** — pick one or more `.xlsx` exports, or
   - **Add folder…** — load every `.xlsx` in a folder at once.
   Already-converted `*_wide.xlsx` files and temporary `~$…` Excel lock files
   are skipped automatically. Use **Remove selected** / **Clear all** to adjust
   the list.

2. **Review & rename variables** — *2. Variables* tab:
   Each detected variable is listed with its type, matrix item, and full
   question text. **Double-click the “Export name” cell** to change the column
   name that variable will receive in the wide file. Press **Enter** to confirm
   or **Esc** to cancel. **Reset names to codes** restores the defaults.
   Renames are keyed to the variable (question code + matrix item), so a name
   you set applies to **every loaded file that uses that variable** — you name
   a variable once for the whole batch. Two columns cannot share the same
   export name: the export stops and lists the clashing names.

3. **Preview** — *3. Preview* tab:
   Click **Refresh preview** to see the exact wide table (up to 50 rows) with
   your current names applied, before writing anything to disk.

4. **Export** — *4. Export* panel:
   - Leave **Combine** unticked → one `<name>_wide.xlsx` per input file, written
     to a folder you choose.
   - Tick **Combine all files into ONE wide sheet** → all submissions from all
     loaded files are stacked into a single workbook (use this when the files
     are the same questionnaire for different classes/sessions).

---

## 5. The output file

Each exported `.xlsx` contains two sheets:

**`wide`** — the data:
- `submission_id`, then `form_id`, `form_title` (meta columns), then
- one column per variable, in questionnaire order (the order in which the
  questions first appear in the input; when files are combined, questions that
  only exist in a later file come after).

**`codebook`** — the legend, one row per variable:

| column          | meaning                                             |
|-----------------|-----------------------------------------------------|
| `variable`      | the stable question code (+ matrix suffix)           |
| `export_name`   | the name you gave it (used as the column in `wide`)  |
| `question_code` | the original question code                           |
| `matrix_row`    | the matrix sub-item text (blank for simple questions)|
| `question_type` | `coded_answer`, `qcm`, or `matrix`                   |
| `page_number`   | page the question appeared on                        |
| `question_text` | the full wording of the question                     |

Keep the codebook with your data — it is the key that maps short column names
back to the actual questions.

---

## 6. How variables are named

- **Simple questions** (`coded_answer`, `qcm`) → the `question_code`
  (e.g. `2db6OGV1sq`).
- **Matrix questions** → `question_code` + the item’s letter/number
  (e.g. `bUfRKycu2X_a`, `bUfRKycu2X_b`), because each matrix row is a separate
  variable. The suffix is read from the item label (`a.`, `b.`, `1)` …). If a
  label has no clean prefix, a short readable slug of the text is used instead,
  and any remaining clash is made unique with a numeric suffix.
- **Multiple detected answers** in one cell (`answer_1`, `answer_2`, `answer_3`)
  are joined with ` | ` — e.g. `Presque tous les jours | Jamais`. This usually
  flags an item the optical reader was unsure about and is worth checking.

Answers are kept **exactly as read**, including `?` / `??` placeholders for
uncertain scans and blanks for no-answer, so you stay in control of any cleaning.
Whole numbers stay whole (`3`, not `3.0`).

---

## 7. Expected input columns

A sheet is treated as the data sheet when it contains at least:
`submission_id`, `question_code`, `question_type`, `matrix_row`, `answer_1`.
The tool also uses, when present: `form_id`, `form_title`, `page_number`,
`question_text`, `answer_2`, `answer_3`. Other columns
(`validation_status`, `confidence`, `review_reason`) are ignored by default.
The correct sheet inside the workbook is detected automatically, so extra
sheets (summaries, notes) do not interfere.

---

## 8. Scripting it (optional, no GUI)

The conversion logic lives in `converter_core.py` and can be used directly:

```python
import converter_core as cc

df = cc.load_long("Test encodage données E01 classe 1.xlsx")
wide, codebook = cc.to_wide(df)
cc.export_wide(wide, codebook, "class1_wide.xlsx")
```

Useful for converting a whole directory in a loop, or wiring the conversion
into an existing pipeline.

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| “No long-format data sheet found” | The workbook has none of the required columns on any sheet — check it is a raw long-format export, not an already-converted file. |
| A file is skipped on **Add folder** | It ends in `_wide.xlsx` (already converted) or is a `~$…` Excel lock file (the workbook is open — close it). |
| SmartScreen blocks the .exe | Unsigned self-built app: **More info → Run anyway**. |
| The browser or antivirus blocks the download | Download it only from this repository's Releases page, then choose **Keep** / allow it. Single-file Python executables are sometimes flagged by mistake. |
| “Several columns would get the same export name” | Two variables (or a variable and `submission_id` / `form_id` / `form_title`) were given the same name in the Variables tab — rename one of them. |
| Column has ` \| ` in a cell | The scanner returned more than one answer for that item — review the original scan. |
| Accents look wrong in a terminal | Display-only; the Excel output stores proper Unicode. |

---

## 10. Files in this repository

| File | Purpose |
|---|---|
| `Questionnaire_Converter.py` | The GUI application source. |
| `converter_core.py` | The conversion logic, GUI-free and importable/scriptable. |
| `requirements.txt` | Python packages needed to run from source. |
| `build.bat` | Builds the standalone `.exe` (see section 11). |
| `tests/` | Unit tests for the conversion logic. |
| `_md_to_pdf.py` | Regenerates `README.pdf` from `README.md`. |
| `README.md` / `README.pdf` | This document. |

The `.exe` itself is not stored in the repository: it is attached to each
[GitHub release](https://github.com/clmntlts/format_conversion/releases).

---

## 11. For maintainers: testing, building and publishing

Run the tests (from the repository folder):

```bash
python -m unittest
```

Build the Windows executable: double-click **`build.bat`** (or run it from a
terminal). It creates a dedicated virtual environment `.venv-build`, installs
`requirements.txt` + PyInstaller, and writes **`dist\Questionnaire_Converter.exe`**.
`build/` and `dist/` are ignored by git.

Regenerate the PDF version of this README after editing it:

```bash
pip install reportlab
python _md_to_pdf.py
```

Publish a new version (with the [GitHub CLI](https://cli.github.com/)), after
committing and pushing the changes:

```bash
gh release create v1.1.0 dist/Questionnaire_Converter.exe --title "v1.1.0" --notes "What changed"
```

The download links in the Quick start always point to the latest release, so
they do not need to be updated.
