# Introduction
hey There I am New Here It's My First Repo 😁

---

## Optimized Excel VBA Worksheet Code

This repo contains **`Sheet1.cls`** — an optimized VBA module for an Excel barcode / inventory workbook.

### How to get the code

**Option 1 — Download the file directly**

1. Open [`Sheet1.cls`](Sheet1.cls) in this repository.
2. Click the **Raw** button (top-right of the file view).
3. Right-click the page → **Save As…** and save it as `Sheet1.cls` on your computer.

**Option 2 — Clone the repository**

```bash
git clone https://github.com/ElliotAbshire01936/Introduction.git
```

The file `Sheet1.cls` will be in the cloned folder.

---

### How to import the code into your Excel workbook

1. Open your Excel workbook (`.xlsm` or `.xlsb` — must be macro-enabled).
2. Press **Alt + F11** to open the Visual Basic Editor (VBE).
3. In the **Project Explorer** on the left, find the entry for **Sheet1** (or whatever your data sheet is named).
4. Right-click the existing **Sheet1** code module → **View Code**.
5. Select **all** existing code (Ctrl + A) and **delete** it.
6. Open the downloaded `Sheet1.cls` file in any text editor (e.g. Notepad), select all (Ctrl + A), copy (Ctrl + C).
7. Paste (Ctrl + V) into the VBE code window.
8. Press **Ctrl + S** to save.  Excel may prompt you to save in macro-enabled format — choose **Yes**.

> **Tip:** Alternatively, you can import via **File → Import File…** in the VBE, but make sure to replace the existing Sheet1 module rather than adding a duplicate.

---

### Required workbook setup

| Requirement | Detail |
|---|---|
| Data sheet | The sheet where this code lives; column **B** holds barcodes, columns **L** and **M** (rows 6–1000) hold manual overrides. |
| **OVERRIDES** sheet | A second sheet named exactly `OVERRIDES`. Column A = barcode key, column B = L-override value, column C = M-override value. |
| Macro-enabled format | Save the workbook as `.xlsm` (Excel Macro-Enabled Workbook) or `.xlsb`. |

---

### What the optimized code does

| Feature | What changed |
|---|---|
| **SelectionChange** | Exits early when the selection is outside rows 1–1000 / columns A–M; recalculates only the used range instead of the whole workbook. |
| **Barcode handling (col B)** | Normalizes each barcode (trim, uppercase, alphanumeric only), stores it as text to preserve leading zeros, and checks for duplicates using a `Scripting.Dictionary` (one scan, O(1) lookups) instead of repeated `CountIf` calls. |
| **Override handling (L6:M1000)** | On edit/paste, writes values to the OVERRIDES sheet keyed by barcode. Uses `Range.Find` instead of `CountIf` per row. After a multi-cell paste, restores default VLOOKUP formulas to rows that have no manual override. |
| **Sort** | `SortDataWithoutHeader` finds the real last data row with `End(xlUp)`, sorts an explicit A–M range, and disables `ScreenUpdating` for speed. |
| **Event safety** | A module-level `mEventGuard` flag prevents re-entrant `Worksheet_Change` fires; `Application.EnableEvents` is always restored in error-handling cleanup labels. |
