"""
Automatic reconciliation table from three inputs in one folder:

  INPUT FOLDER (drop these here):
  1) BKERROUT*     — text report with BKDATE, BANKNAME, DOCS, BANK_AMT, PROCESSED_AMT
  2) SQL.CSV       — OPR ID, DOCS, AMOUNT (or similar header)
  3) opr_mapping.json (optional) — OPR ID → Bank display name. Defaults to bundled mapping.

  OUTPUT:
  reconciliation_output.csv (and reconciliation_output.xlsx if openpyxl is installed)

  Usage:
    python bank_reconciliation_table.py
    python bank_reconciliation_table.py --input "C:\\path\\to\\my_inputs"
    python bank_reconciliation_table.py -i ./input
    python bank_reconciliation_table.py -i reconciliation_input --gui
      # Opens Tkinter: checkboxes + black initials box + Submit → writes CSV/XLSX
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = SCRIPT_DIR / "reconciliation_input"


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().upper())


def format_amount_currency(value) -> str:
    """
    Format amounts as $12,34,567.89 (Indian grouping: last 3 digits, then pairs)
    with $ prefix. Empty/invalid returns "".
    """
    if value is None or value == "":
        return ""
    try:
        n = float(value)
    except (TypeError, ValueError):
        return str(value)
    negative = n < 0
    n = abs(n)
    int_part = int(n)
    frac = round((n - int_part) * 100)
    if frac >= 100:
        int_part += 1
        frac = 0
    s = str(int_part)
    # Indian grouping: ... 12,34,567
    if len(s) <= 3:
        grouped = s
    else:
        parts = [s[-3:]]
        i = len(s) - 3
        while i > 0:
            start = max(0, i - 2)
            parts.insert(0, s[start:i])
            i = start
        grouped = ",".join(parts)
    sign = "-" if negative else ""
    return f"{sign}${grouped}.{frac:02d}"


def load_opr_mapping(input_dir: Path) -> dict[str, str]:
    """OPR ID -> display bank name."""
    p = input_dir / "opr_mapping.json"
    if p.is_file():
        with p.open(encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {k.strip().upper(): v for k, v in data.items()}
    # Bundled fallback
    bundled = SCRIPT_DIR / "opr_mapping.json"
    if bundled.is_file():
        with bundled.open(encoding="utf-8") as f:
            data = json.load(f)
        return {k.strip().upper(): v for k, v in data.items()}
    return {}


def find_bkerrou(input_dir: Path) -> Path | None:
    """First file whose name contains BKERROUT (case-insensitive)."""
    for p in sorted(input_dir.iterdir()):
        if p.is_file() and "bkerrou" in p.name.lower():
            return p
    return None


def find_sql_csv(input_dir: Path) -> Path | None:
    """Prefer SQL.CSV; else first CSV that looks like OPR/DOCS/AMOUNT."""
    candidates = []
    for p in sorted(input_dir.iterdir()):
        if not p.is_file() or not p.suffix.lower() == ".csv":
            continue
        if "reconciliation_output" in p.name.lower():
            continue
        candidates.append(p)
    for p in candidates:
        if p.name.upper().startswith("SQL"):
            return p
    for p in candidates:
        try:
            head = p.read_text(encoding="utf-8-sig", errors="replace")[:2000]
            blob = head.upper()
            if "OPR" in blob and ("DOCS" in blob or "AMOUNT" in blob):
                return p
        except OSError:
            continue
    return candidates[0] if candidates else None


def parse_bkerrou(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8", errors="replace")
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("---"):
            continue
        if "BKDATE" in line and "BANKNAME" in line:
            continue
        parts = re.split(r"\s{2,}", line)
        if len(parts) < 6:
            continue
        try:
            bdate, bank = parts[0], parts[1].strip()
            docs = int(parts[2].replace(",", ""))
            bank_amt = float(parts[3].replace(",", ""))
            proc_amt = float(parts[4].replace(",", ""))
            err_s = parts[5].replace(",", "").strip()
            err_amt = float(err_s) if err_s and err_s != "." else 0.0
            if err_s in (".00", ".0", "."):
                err_amt = 0.0
        except (ValueError, IndexError):
            continue
        rows.append({
            "BKDATE": bdate,
            "BANKNAME": bank,
            "DOCS": docs,
            "BANK_AMT": bank_amt,
            "PROCESSED_AMT": proc_amt,
            "ERROR_AMT": err_amt,
        })
    return rows


def parse_sql_csv(path: Path, opr_to_bank: dict[str, str]) -> dict[str, dict]:
    """Normalized bank key -> {OPR_ID, DOCS, AMOUNT}."""
    out = {}
    with path.open(newline="", encoding="utf-8-sig", errors="replace") as f:
        sample = f.read(4096)
        f.seek(0)
        # Sniff delimiter
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        f.seek(0)
        r = csv.DictReader(f, dialect=dialect)
        if not r.fieldnames:
            return out
        # Map flexible column names
        fn = [x.strip() if x else "" for x in r.fieldnames]
        opr_key = None
        for c in fn:
            cu = c.upper()
            if "OPR" in cu and "ID" in cu:
                opr_key = c
                break
        if not opr_key:
            opr_key = fn[0]
        docs_key = next((c for c in fn if "DOC" in c.upper()), None)
        amt_key = next((c for c in fn if "AMOUNT" in c.upper() or "AMT" in c.upper()), None)
        if not docs_key:
            docs_key = fn[1] if len(fn) > 1 else None
        if not amt_key:
            amt_key = fn[2] if len(fn) > 2 else None

        for row in r:
            opr = (row.get(opr_key) or "").strip()
            if not opr:
                continue
            bank = opr_to_bank.get(opr.upper()) or opr_to_bank.get(opr)
            if not bank:
                continue
            try:
                docs = int(float((row.get(docs_key) or "0").replace(",", "")))
                amt = float(str(row.get(amt_key) or "0").replace(",", ""))
            except (ValueError, TypeError):
                continue
            out[norm(bank)] = {"OPR_ID": opr, "DOCS": docs, "AMOUNT": amt}
    return out


def build_table(
    bkerrou_rows: list[dict],
    sql_by_bank: dict[str, dict],
    opr_to_bank: dict[str, str],
) -> list[dict]:
    mail_by_bank = {}
    err2_by_bank = {}
    for r in bkerrou_rows:
        key = norm(r["BANKNAME"])
        mail_by_bank[key] = {
            "BKDATE": r["BKDATE"],
            "BANKNAME": r["BANKNAME"],
            "DOCS": r["DOCS"],
            "BANK_AMT": r["BANK_AMT"],
        }
        err2_by_bank[key] = {
            "DOCS": r["DOCS"],
            "PROCESSED_AMT": r["PROCESSED_AMT"],
        }

    all_keys = set(mail_by_bank) | set(sql_by_bank)
    ordered = []
    seen = set()
    for r in bkerrou_rows:
        k = norm(r["BANKNAME"])
        if k not in seen and k in all_keys:
            ordered.append(k)
            seen.add(k)
    for k in sorted(all_keys - seen):
        ordered.append(k)

    out_rows = []
    for key in ordered:
        mail = mail_by_bank.get(key, {})
        sql = sql_by_bank.get(key, {})
        err2 = err2_by_bank.get(key, {})
        display = mail.get("BANKNAME") or next(
            (opr_to_bank[o] for o, n in opr_to_bank.items() if norm(str(n)) == key),
            key,
        )
        out_rows.append({
            "Likely_entry_Bank_Name": display,
            "Mail_Count": mail.get("DOCS", ""),
            "Mail_Amount": mail.get("BANK_AMT", ""),
            "SQL_Count": sql.get("DOCS", ""),
            "SQL_Amount": sql.get("AMOUNT", ""),
            "Err2_Count": err2.get("DOCS", ""),
            "Err2_Amount": err2.get("PROCESSED_AMT", ""),
            "Verified": "",  # manual / formula
            "BKDATE": mail.get("BKDATE", ""),
        })
    return out_rows


def write_csv(rows: list[dict], out_path: Path) -> None:
    # Single-row headers matching template sections (CSV cannot merge cells)
    fieldnames = [
        "Bank Name",
        "Bank Mail's Data - Count",
        "Bank Mail's Data - Amount",
        "SQL Query output - Count",
        "SQL Query output - Amount",
        "Bank Err 2 - Count",
        "Bank Err 2 - Amount",
        "Verified",
        "BKDATE",
    ]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(fieldnames)
        for r in rows:
            w.writerow([
                r.get("Likely_entry_Bank_Name", ""),
                r.get("Mail_Count", ""),
                format_amount_currency(r.get("Mail_Amount", "")),
                r.get("SQL_Count", ""),
                format_amount_currency(r.get("SQL_Amount", "")),
                r.get("Err2_Count", ""),
                format_amount_currency(r.get("Err2_Amount", "")),
                r.get("Verified", ""),
                r.get("BKDATE", ""),
            ])


def write_xlsx(rows: list[dict], out_path: Path) -> bool:
    """Write Excel matching template: merged header row + colored sections + Verified column."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    except ImportError:
        return False

    # Template colors (approximate to typical Excel fills)
    YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Bank Name, Bank Err 2
    BLUE = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")   # Bank Mail's Data, Count/Amount sub
    ORANGE = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")  # SQL Query output
    GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Verified
    WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    # Thick borders on all cells for clear column/row grid
    _thick = Side(style="thick", color="000000")
    all_border = Border(left=_thick, right=_thick, top=_thick, bottom=_thick)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_font = Font(bold=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    # Row 1: group headers (merged)
    # A1:A2 = Bank Name
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.value = "Bank Name"
    c.fill = YELLOW
    c.font = header_font
    c.alignment = center
    c.border = all_border
    ws["A2"].border = all_border

    # B1:C1 = Bank Mail's Data
    ws.merge_cells("B1:C1")
    ws["B1"].value = "Bank Mail's Data"
    ws["B1"].fill = BLUE
    ws["B1"].font = header_font
    ws["B1"].alignment = center
    ws["B1"].border = all_border
    ws["C1"].border = all_border

    # D1:E1 = SQL Query output
    ws.merge_cells("D1:E1")
    ws["D1"].value = "SQL Query output"
    ws["D1"].fill = ORANGE
    ws["D1"].font = header_font
    ws["D1"].alignment = center

    # F1:G1 = Bank Err 2
    ws.merge_cells("F1:G1")
    ws["F1"].value = "Bank Err 2"
    ws["F1"].fill = YELLOW
    ws["F1"].font = header_font
    ws["F1"].alignment = center

    # H1:H2 = Verified
    ws.merge_cells("H1:H2")
    ws["H1"].value = "Verified"
    ws["H1"].fill = GREEN
    ws["H1"].font = header_font
    ws["H1"].alignment = center

    # Row 2: sub-headers Count / Amount
    sub_headers = [
        ("B2", "Count", BLUE),
        ("C2", "Amount", BLUE),
        ("D2", "Count", ORANGE),
        ("E2", "Amount", ORANGE),
        ("F2", "Count", BLUE),  # light blue sub like template
        ("G2", "Amount", BLUE),
    ]
    for ref, text, fill in sub_headers:
        cell = ws[ref]
        cell.value = text
        cell.fill = fill
        cell.font = header_font
        cell.alignment = center
        cell.border = all_border

    # Borders for merged row1 cells
    for ref in ("B1", "C1", "D1", "E1", "F1", "G1"):
        ws[ref].border = all_border
    for col in range(1, 9):
        ws.cell(row=2, column=col).border = all_border
    ws["H2"].border = all_border

    # Data rows start at row 3 — amount columns formatted as $12,34,567.89
    data_start = 3
    for i, row in enumerate(rows):
        r = data_start + i
        values = [
            row.get("Likely_entry_Bank_Name", ""),
            row.get("Mail_Count", ""),
            format_amount_currency(row.get("Mail_Amount", "")),
            row.get("SQL_Count", ""),
            format_amount_currency(row.get("SQL_Amount", "")),
            row.get("Err2_Count", ""),
            format_amount_currency(row.get("Err2_Amount", "")),
            row.get("Verified", ""),
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.border = all_border
            if col == 1:
                cell.alignment = Alignment(vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Column widths (A–H only; BKDATE is not written to Excel)
    ws.column_dimensions["A"].width = 28
    for col_letter in ("B", "C", "D", "E", "F", "G"):
        ws.column_dimensions[col_letter].width = 14
    ws.column_dimensions["H"].width = 12

    wb.save(out_path)
    return True


def prompt_bank_selection(rows: list[dict]) -> tuple[list[dict], str] | tuple[None, None]:
    """
    Tkinter window integrated with reconciliation: checkboxes per bank + verifier
    initials in a black entry box. Submit writes CSV/XLSX with selected rows.
    Returns (filtered_rows, initials) or (None, None) if cancelled.
    """
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        sys.stderr.write("Tkinter not available; install python with tk support.\n")
        return None, None

    if not rows:
        return [], ""

    result_rows: list[dict] = []
    result_initials = ""
    cancelled = {"v": True}

    root = tk.Tk()
    root.title("Bank reconciliation — Select & Submit")
    root.geometry("540x520")
    root.minsize(420, 340)
    root.configure(bg="#f0f0f0")

    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        main,
        text="Select banks to include, then enter verifier initials and click Submit.",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor=tk.W)
    ttk.Label(
        main,
        text="Unchecked banks are excluded from the exported table.",
        font=("Segoe UI", 9),
        foreground="#555",
    ).pack(anchor=tk.W, pady=(0, 8))

    # Scrollable checkbox area
    list_outer = ttk.Frame(main)
    list_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
    canvas = tk.Canvas(list_outer, highlightthickness=1, highlightbackground="#ccc", bg="white")
    scroll = ttk.Scrollbar(list_outer, orient=tk.VERTICAL, command=canvas.yview)
    cb_frame = ttk.Frame(canvas)
    cb_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    _win_id = canvas.create_window((0, 0), window=cb_frame, anchor=tk.NW)
    canvas.configure(yscrollcommand=scroll.set)

    def _canvas_configure(event):
        canvas.itemconfigure(_win_id, width=max(event.width - 4, 100))

    canvas.bind("<Configure>", _canvas_configure)

    # One BooleanVar per row (duplicate bank names allowed)
    check_vars: list[tk.BooleanVar] = []
    for i, row in enumerate(rows):
        name = str(row.get("Likely_entry_Bank_Name", f"Row {i+1}"))
        var = tk.BooleanVar(value=True)
        check_vars.append(var)
        ttk.Checkbutton(cb_frame, text=name, variable=var).pack(anchor=tk.W, padx=8, pady=3)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

    # Verifier initials — black box (tk.Entry for full bg/fg control)
    ttk.Label(main, text="Verifier initials (black box):", font=("Segoe UI", 9, "bold")).pack(
        anchor=tk.W, pady=(4, 4)
    )
    black_box = tk.Frame(main, bg="black", padx=10, pady=10)
    black_box.pack(fill=tk.X, pady=(0, 6))
    initials_entry = tk.Entry(
        black_box,
        bg="black",
        fg="white",
        insertbackground="white",
        relief=tk.FLAT,
        font=("Consolas", 14),
        justify=tk.CENTER,
    )
    initials_entry.pack(fill=tk.X, ipady=8)
    initials_entry.focus_set()
    ttk.Label(
        main,
        text="Typed initials are copied into the Verified column for each included bank.",
        font=("Segoe UI", 8),
        foreground="gray",
    ).pack(anchor=tk.W)

    btn_frame = ttk.Frame(main)
    btn_frame.pack(fill=tk.X, pady=(14, 0))

    from tkinter import messagebox

    def on_submit():
        nonlocal result_rows, result_initials
        initials = initials_entry.get().strip()
        selected = []
        for i, row in enumerate(rows):
            if i >= len(check_vars) or not check_vars[i].get():
                continue
            r = dict(row)
            r["Verified"] = initials
            selected.append(r)
        if not selected:
            messagebox.showwarning("No selection", "Select at least one bank.", parent=root)
            return
        result_rows = selected
        result_initials = initials
        cancelled["v"] = False
        root.destroy()

    def on_cancel():
        root.destroy()

    # Submit = primary action (export); Cancel closes without writing
    ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=(10, 0))
    submit_btn = tk.Button(
        btn_frame,
        text="Submit",
        command=on_submit,
        bg="#2563eb",
        fg="white",
        activebackground="#1d4ed8",
        activeforeground="white",
        font=("Segoe UI", 11, "bold"),
        padx=24,
        pady=8,
        cursor="hand2",
    )
    submit_btn.pack(side=tk.RIGHT)
    root.bind("<Return>", lambda e: on_submit())
    root.bind("<Escape>", lambda e: on_cancel())

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()

    if cancelled["v"]:
        return None, None
    return result_rows, result_initials


def run(input_dir: Path, use_gui: bool = False) -> Path | None:
    input_dir = input_dir.resolve()
    if not input_dir.is_dir():
        sys.stderr.write(f"Input folder does not exist: {input_dir}\n")
        sys.exit(1)

    opr_to_bank = load_opr_mapping(input_dir)
    if not opr_to_bank:
        sys.stderr.write(
            "No OPR mapping found. Add opr_mapping.json to the input folder or script folder.\n"
        )
        sys.exit(1)

    bkerrou_path = find_bkerrou(input_dir)
    sql_path = find_sql_csv(input_dir)
    if not bkerrou_path:
        sys.stderr.write(
            f"No BKERROUT file found in {input_dir}. "
            "Place a file whose name contains BKERROUT there.\n"
        )
        sys.exit(1)
    if not sql_path:
        sys.stderr.write(
            f"No SQL CSV found in {input_dir}. Place SQL.CSV (OPR ID, DOCS, AMOUNT) there.\n"
        )
        sys.exit(1)

    bkerrou_rows = parse_bkerrou(bkerrou_path)
    sql_by_bank = parse_sql_csv(sql_path, opr_to_bank)
    rows = build_table(bkerrou_rows, sql_by_bank, opr_to_bank)

    if use_gui:
        rows, _ = prompt_bank_selection(rows)
        if rows is None:
            return None  # cancelled

    out_csv = input_dir / "reconciliation_output.csv"
    write_csv(rows, out_csv)
    out_xlsx = input_dir / "reconciliation_output.xlsx"
    if write_xlsx(rows, out_xlsx):
        return out_csv  # caller can print both
    return out_csv


def main():
    ap = argparse.ArgumentParser(
        description="Build reconciliation table from BKERROUT + SQL.CSV in one folder."
    )
    ap.add_argument(
        "-i", "--input",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=f"Folder containing BKERROUT file + SQL.CSV (default: {DEFAULT_INPUT_DIR})",
    )
    ap.add_argument("-q", "--quiet", action="store_true", help="Only print output path")
    ap.add_argument(
        "--gui",
        action="store_true",
        help="Open Tkinter window to select banks and enter verifier initials before export",
    )
    args = ap.parse_args()

    input_dir = args.input
    if not input_dir.is_absolute():
        input_dir = (Path.cwd() / input_dir).resolve()

    out_csv = run(input_dir, use_gui=args.gui)
    if out_csv is None and args.gui:
        if not args.quiet:
            print("Cancelled — no file written.")
        return
    out_xlsx = input_dir / "reconciliation_output.xlsx"
    if not args.quiet:
        print(f"Input folder: {input_dir}")
        print(f"Wrote: {out_csv}")
        if out_xlsx.is_file():
            print(f"Wrote: {out_xlsx}")
    else:
        print(out_csv)


if __name__ == "__main__":
    main()
