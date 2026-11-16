# Questionnaire Long → Wide Converter

A desktop tool that converts optically-read questionnaire exports from
**long format** (one row per question / matrix item) to **wide format**
(one row per `submission_id`, one column per variable), ready for analysis
in Excel, SPSS, R, JASP, or Python.

The questionnaires are fully modular, so the tool does not assume a fixed set
of questions: it reads whatever question codes and matrix items are actually
present in each file and builds the variables from them.

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

- **Using the .exe:** nothing — it is self-contained and runs on Windows
  without Python installed.
- **Running from source:** Python 3.10+ with `pandas` and `openpyxl`.
  Install with:

  ```bash
  pip install pandas openpyxl
  ```

  `tkinter` (the GUI toolkit) ships with the standard Python installer on Windows.

---

## 3. Running the tool

- **Easiest — the executable:** double-click
  **`dist/Questionnaire_Converter.exe`**.
  The first launch takes a few seconds while it unpacks; later launches are faster.
- **From source:**

  ```bash
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
   Renames are keyed to the question code, so a name you set applies to **every
   loaded file that uses that code** — you name a variable once for the whole batch.

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
- one column per variable, in questionnaire order (by page, then original order).

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
| Column has ` | ` in a cell | The scanner returned more than one answer for that item — review the original scan. |
| Accents look wrong in a terminal | Display-only; the Excel output stores proper Unicode. |

---

## 10. Files in this folder

| File | Purpose |
|---|---|
| `dist/Questionnaire_Converter.exe` | Standalone Windows executable (no Python needed). |
| `Questionnaire_Converter.py` | The GUI application source. |
| `converter_core.py` | The conversion logic, GUI-free and importable/scriptable. |
| `README.md` / `README.pdf` | This document. |
