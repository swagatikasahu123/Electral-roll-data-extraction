

import os
import re
import zipfile
import tempfile
import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------- Regex / heuristics ----------
EPIC_REGEX = re.compile(r'\b[A-Z0-9]{5,}\b')
AGE_REGEX = re.compile(r'\b([1-9][0-9]?)\b')
GENDER_REGEX = re.compile(r'\b(M|F|m|f|पुरुष|महिला|स्त्री)\b', re.IGNORECASE)
RELATION_TYPE_MAP = {
    'S/O': 'Father', 'SON OF': 'Father', 'FATHER': 'Father',
    'W/O': 'Husband', 'WIFE OF': 'Husband', 'HUSBAND': 'Husband',
    'पिता': 'Father', 'पति': 'Husband'
}
VIDHAN_REGEX = re.compile(r'(\d+)\s*-\s*([^\n]+)')
BOOTH_NO_REGEX = re.compile(r'भाग\s*[:\-]?\s*(\d+)', re.IGNORECASE)
ST_CODE_REGEX = re.compile(r'\b(S\d{2})\b', re.IGNORECASE)

# ---------- Placeholder mappings (extend these) ----------
ST_CODE_TO_STATE = {
    "S04": "Bihar"
}
AC_NO_TO_NAME = {
    "240": "Sikandra"
}

# ---------- Utilities ----------
def maybe_unzip(path):
    if path.lower().endswith(".zip") and os.path.isfile(path):
        tempdir = tempfile.mkdtemp(prefix="eroll_")
        with zipfile.ZipFile(path, "r") as z:
            z.extractall(tempdir)
        return tempdir
    return path

def normalize_gender(g):
    if not g:
        return None
    g = g.strip().upper()
    if g in ["M", "पुरुष"]:
        return "M"
    if g in ["F", "महिला", "स्त्री"]:
        return "F"
    return g

def extract_header_metadata(first_page):
    text = first_page.extract_text() or ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    state_name = ""
    vidhan_name = ""
    vidhan_number = ""
    booth_name = ""
    booth_number = ""

    # ST_CODE
    for line in lines:
        m = ST_CODE_REGEX.search(line)
        if m:
            st_code = m.group(1)
            state_name = ST_CODE_TO_STATE.get(st_code, "")
            break

    # Vidhan Sabha (AC) number & name
    for line in lines:
        m = VIDHAN_REGEX.search(line)
        if m:
            vidhan_number = m.group(1).strip()
            vidhan_name = AC_NO_TO_NAME.get(vidhan_number, m.group(2).strip())
            break

    # Booth number
    for line in lines:
        m = BOOTH_NO_REGEX.search(line)
        if m:
            booth_number = m.group(1).strip()
            break

    # Booth name (heuristic near मतदान केंद्र)
    for i, line in enumerate(lines):
        if "मतदान" in line or "केंद्र" in line:
            for nxt in lines[i : i + 3]:
                m = re.search(r'\d+\s*-\s*(.+)', nxt)
                if m:
                    booth_name = m.group(1).strip()
                    break
            if booth_name:
                break

    # Fallbacks
    if not state_name:
        state_name = "UnknownState"
    if not vidhan_name:
        vidhan_name = "UnknownVidhanSabha"
    if not vidhan_number:
        vidhan_number = ""
    if not booth_name:
        booth_name = "UnknownBooth"
    if not booth_number:
        booth_number = ""

    return {
        "State Name": state_name,
        "Vidhan sabha Name": vidhan_name,
        "Vidhan sabha Number": vidhan_number,
        "Booth Name": booth_name,
        "Booth Number": booth_number
    }

def parse_voter_block(text, header_meta, source_pdf):
    norm = re.sub(r'([A-Z0-9])\s+([A-Z0-9])', r'\1\2', text)  # collapse split EPIC
    epic_m = EPIC_REGEX.search(norm)
    epic_no = epic_m.group(0) if epic_m else None

    gender_m = GENDER_REGEX.search(norm)
    gender = normalize_gender(gender_m.group(1)) if gender_m else None

    age = None
    for a in re.findall(r'\b([1-9][0-9]?)\b', norm):
        if 17 < int(a) < 121:
            age = a
            break

    serial = None
    mserial = re.match(r'^\s*(\d+)', norm)
    if mserial:
        serial = mserial.group(1)

    relation_type = None
    relation_name = None
    for key, val in RELATION_TYPE_MAP.items():
        if key.upper() in norm.upper():
            relation_type = val
            # try capture following name token(s)
            m = re.search(rf'{re.escape(key)}\s+([A-Za-z\u0900-\u097F]+)', norm, re.IGNORECASE)
            if m:
                relation_name = m.group(1)
            break
    # fallback Hindi explicit
    if not relation_type:
        if re.search(r'पिता', norm):
            relation_type = "Father"
        elif re.search(r'पति', norm):
            relation_type = "Husband"

    # Voter name: English-first fallback to Hindi
    fm_name_en = None
    last_en = None
    en_names = re.findall(r'\b([A-Z][a-z]+)\b', norm)
    if len(en_names) >= 2:
        fm_name_en = en_names[0]
        last_en = en_names[1]
    elif len(en_names) == 1:
        fm_name_en = en_names[0]

    fm_name_hi = None
    last_hi = None
    hi_names = re.findall(r'[\u0900-\u097F]+', norm)
    if len(hi_names) >= 2:
        fm_name_hi = hi_names[0]
        last_hi = hi_names[1]
    elif len(hi_names) == 1:
        fm_name_hi = hi_names[0]

    # House number
    house = None
    mh = re.search(r'(?:H\.?\s*No\.?:?\s*|मकान\s*सखंया\s*[:：]?\s*)(\w+)', norm, re.IGNORECASE)
    if mh:
        house = mh.group(1)

    # Build voter's name
    voter_name = None
    if fm_name_en and last_en:
        voter_name = f"{fm_name_en} {last_en}"
    elif fm_name_hi and last_hi:
        voter_name = f"{fm_name_hi} {last_hi}"
    elif fm_name_en:
        voter_name = fm_name_en
    elif fm_name_hi:
        voter_name = fm_name_hi

    # Compose final structured dictionary
    row = {
        "State Name": header_meta.get("State Name"),
        "Vidhan sabha Name & Number": f"{header_meta.get('Vidhan sabha Name','')} {header_meta.get('Vidhan sabha Number','')}".strip(),
        "Booth Name & Number": f"{header_meta.get('Booth Name','')} {header_meta.get('Booth Number','')}".strip(),
        "Voter's Serial Number": serial,
        "Voter's Name": voter_name,
        "Voter ID (EPIC Number)": epic_no,
        "Relation's Name (Father's or Husband's)": relation_name,
        "Relation Type (Father or Husband)": relation_type,
        "House Number": house,
        "Age": age,
        "Gender": gender,
        "Source PDF": os.path.basename(source_pdf)
    }
    return row

def extract_from_pdf(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        header_meta = {}
        if pdf.pages:
            header_meta = extract_header_metadata(pdf.pages[0])
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if not txt.strip():
                continue
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            block_accum = []
            for line in lines:
                if EPIC_REGEX.search(line):
                    block_accum.append(line)
                else:
                    if block_accum:
                        block_accum[-1] = block_accum[-1] + " " + line
            for blk in block_accum:
                row = parse_voter_block(blk, header_meta, pdf_path)
                if row:
                    rows.append(row)
    return rows

# ---------- Summaries & output ----------
def build_summary(df):
    def missing_critical(r):
        return pd.isna(r["Voter ID (EPIC Number)"]) or pd.isna(r["Gender"]) or pd.isna(r["Age"])
    summary = (
        df.groupby("Source PDF")
          .agg(
              Voter_Count=("Voter's Name", "count"),
              Missing_Critical=("Source PDF", lambda x: df.loc[x.index].apply(missing_critical, axis=1).sum()),
              Unique_EPICs=("Voter ID (EPIC Number)", lambda s: s.dropna().nunique())
          )
          .reset_index()
    )
    return summary

def run_gui():
    root = tk.Tk()
    root.title("Electoral Roll Data Extraction Tool")

    input_var = tk.StringVar()
    output_var = tk.StringVar()

    def choose_input():
        p = filedialog.askopenfilename(title="Select ZIP or PDF") or filedialog.askdirectory(title="Or select folder")
        input_var.set(p)

    def choose_output():
        p = filedialog.askdirectory(title="Select output folder")
        output_var.set(p)

    def execute():
        inp = input_var.get()
        out = output_var.get()
        if not inp or not out:
            messagebox.showerror("Error", "Both input and output must be provided.")
            return
        try:
            resolved = maybe_unzip(inp)
            pdfs = []
            if os.path.isdir(resolved):
                for root_dir, _, files in os.walk(resolved):
                    for f in files:
                        if f.lower().endswith(".pdf"):
                            pdfs.append(os.path.join(root_dir, f))
            elif os.path.isfile(resolved) and resolved.lower().endswith(".pdf"):
                pdfs = [resolved]
            else:
                raise ValueError("Input must be PDF file/folder/ZIP.")

            all_rows = []
            for pdf in sorted(pdfs):
                print(f"[=] Parsing {pdf}")
                extracted = extract_from_pdf(pdf)
                all_rows.extend(extracted)

            if not all_rows:
                messagebox.showwarning("No data", "No voter rows were extracted.")
                return

            df = pd.DataFrame(all_rows)
            # reorder columns as requested (drop Source PDF in cleaned output if undesired)
            cleaned_cols = [
                "State Name",
                "Vidhan sabha Name & Number",
                "Booth Name & Number",
                "Voter's Serial Number",
                "Voter's Name",
                "Voter ID (EPIC Number)",
                "Relation's Name (Father's or Husband's)",
                "Relation Type (Father or Husband)",
                "House Number",
                "Age",
                "Gender"
            ]
            cleaned = df.reindex(columns=cleaned_cols + ["Source PDF"])  # keep source for summary

            # Save cleaned voter-level
            os.makedirs(out, exist_ok=True)
            cleaned_path = os.path.join(out, "final_cleaned_output.xlsx")
            cleaned.to_excel(cleaned_path, index=False)

            # Summary per PDF
            summary = build_summary(cleaned)
            summary_path = os.path.join(out, "per_pdf_summary.xlsx")
            summary.to_excel(summary_path, index=False)

            messagebox.showinfo(
                "Done",
                f"Extraction complete.\nVoter-level: {cleaned_path}\nSummary: {summary_path}"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # GUI layout
    tk.Label(root, text="Input ZIP / Folder / PDF:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(root, textvariable=input_var, width=50).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=choose_input).grid(row=0, column=2, padx=5)

    tk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(root, textvariable=output_var, width=50).grid(row=1, column=1)
    tk.Button(root, text="Browse", command=choose_output).grid(row=1, column=2, padx=5)

    tk.Button(root, text="Run Extraction", command=execute, bg="#2563EB", fg="white", padx=12, pady=6).grid(row=2, column=1, pady=12)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
