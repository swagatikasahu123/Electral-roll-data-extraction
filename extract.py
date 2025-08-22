

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Electoral Roll Data Extraction Tool
- Supports CLI (argparse) and GUI (tkinter)
- Parses multi-column Hindi electoral PDFs using pdfplumber with regex heuristics
- Consolidates to a single Excel

Usage (CLI):
  python extract.py --input "C:\folder\rolls" --output "C:\out\final_output.xlsx"

GUI:
  python extract.py --gui
"""
import os, re, sys, argparse, glob
from typing import List
import pandas as pd

# ---------- Normalization ----------
def norm(s: str) -> str:
    if not isinstance(s, str):
        return s
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"[ \xa0]+", " ", s)
    return s.replace("\u200c","").replace("\u200b","").strip()

def norm_gender(g: str) -> str:
    g = norm(g)
    if "पुरुष" in g or "परुष" in g or "परुुष" in g:
        return "Male"
    if "महिला" in g or "मिहला" in g:
        return "Female"
    return g

# ---------- Regex ----------
serial_epic_re = re.compile(r"^\s*(\d{1,4})\s+([A-Z]{2,4}[/A-Z0-9-]*\d{3,})\b")
stop_tokens = r"(?:िनवार्चक का नाम|पिता का नाम|पित का नाम|पति का नाम|पत्नी का नाम|माता का नाम|मकान|उम्र|लिंग|\n)"
name_re = re.compile(r"िनवार्चक का नाम\s*[:：]\s*(.+?)(?=\s*" + stop_tokens + r")")
rel_re = re.compile(r"(पिता|पित|पति|पत्नी|माता)\s*(?:का\s*नाम)?\s*[:：]?\s*([^\n]+)")   # FIXED
house_re = re.compile(r"मकान\s*स(?:ं|ंख्या|खं|ख्या)\s*[:：]?\s*(.+?)(?=\s*" + stop_tokens + r")")
age_gender_re = re.compile(r"उम्र\s*[:：]?\s*(\d{1,3}).*?(?:लि?ं?ग)\s*[:：]?\s*([^\n]+)")

# ---------- Helpers ----------
def split_blocks(text: str) -> List[str]:
    lines = text.splitlines()
    blocks = []
    current = []
    for line in lines:
        if serial_epic_re.match(line):
            if current:
                blocks.append("\n".join(current[:20]))
                current = []
        current.append(line)
    if current:
        blocks.append("\n".join(current[:20]))
    return blocks

# ---------- Parser ----------
def parse_pdf(path: str) -> pd.DataFrame:
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            pages_text = [norm(p.extract_text(x_tolerance=2, y_tolerance=2) or "") for p in pdf.pages]
    except Exception:
        import PyPDF2
        reader = PyPDF2.PdfReader(open(path, "rb"))
        pages_text = [norm(page.extract_text() or "") for page in reader.pages]

    header_text = "\n".join(pages_text[:3])

    # State
    state = "Bihar" if "बिहार" in header_text or "िबहार" in header_text else "Unknown"

    # Vidhan Sabha
    m_vs = re.search(r"िवधानसभा[^:]*[:：]\s*([0-9]+)\s*-\s*([^\n(]+)", header_text)
    vs_num, vs_name = (m_vs.group(1).strip(), norm(m_vs.group(2))) if m_vs else ("", "")

    # Booth (Universal Regex)
    booth_num, booth_name = "", ""
    header_lines = header_text.splitlines()
    for idx, line in enumerate(header_lines):
        if re.search(r"(भाग\s*सं|भाग\s*संख्या|मतदान\s*कें?द्र|Booth\s*No)", line):
            # Try same line
            m = re.search(r"([0-9]+)\s*[-–]\s*([^\n]+)", line)
            if m:
                booth_num, booth_name = m.group(1).strip(), norm(m.group(2))
                break
            # Try next line
            if idx + 1 < len(header_lines):
                m2 = re.search(r"([0-9]+)\s*[-–]\s*([^\n]+)", header_lines[idx+1])
                if m2:
                    booth_num, booth_name = m2.group(1).strip(), norm(m2.group(2))
                    break

    rows = []
    for page_text in pages_text:
        for block in split_blocks(page_text):
            lines = [norm(x) for x in block.splitlines() if x.strip()]
            if not lines: 
                continue
            m_head = serial_epic_re.match(lines[0])
            if not m_head: 
                continue
            serial, epic = m_head.group(1), m_head.group(2)
            body = "\n".join(lines[1:])
            name = rel_name = house = age = gender = ""
            rel_type = ""

            # Name
            mn = name_re.search(body) 
            if mn: 
                name = norm(mn.group(1))

            # Relation (FIXED)
            mr = rel_re.search(body)
            if mr:
                rel_name = norm(mr.group(2))
                rk = mr.group(1)
                if "पिता" in rk or "पित" in rk:
                    rel_type = "Father"
                elif "पति" in rk:
                    rel_type = "Husband"
                elif "पत्नी" in rk:
                    rel_type = "Wife"
                elif "माता" in rk:
                    rel_type = "Mother"

            # House
            mh = house_re.search(body)
            if mh:
                house = norm(mh.group(1))
                house = house.split("फोटो")[0].strip()

            # Age/Gender
            mag = age_gender_re.search(body)
            if mag:
                age = mag.group(1)
                gender = norm_gender(mag.group(2))

            rows.append({
                "State Name": state,
                "Vidhan sabha Name & Number": f"{vs_num}-{vs_name}" if vs_num else "",
                "Booth Name & Number": f"{booth_num}-{booth_name}" if booth_num else "Not Found",
                "Voter's Serial Number": serial,
                "Voter's Name": name,
                "Voter ID (EPIC Number)": epic,
                "Relation's Name": rel_name,
                "Relation Type": rel_type,
                "House Number": re.findall(r"\d+", house)[-1] if re.findall(r"\d+", house) else house,
                "Age": age,
                "Gender": gender
            })
    return pd.DataFrame(rows)

# ---------- Helpers for multiple files ----------
def gather_pdfs(input_path: str) -> List[str]:
    if os.path.isdir(input_path):
        pdfs = sorted(glob.glob(os.path.join(input_path, "*.pdf")))
    elif os.path.isfile(input_path) and input_path.lower().endswith(".pdf"):
        pdfs = [input_path]
    else:
        pdfs = []
    return pdfs

def run_cli(args):
    pdfs = gather_pdfs(args.input)
    if not pdfs:
        print("No PDFs found in input path.")
        sys.exit(1)
    frames = []
    for p in pdfs:
        df = parse_pdf(p)
        frames.append(df)
    out_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    os.makedirs(os.path.dirname(args.output), exist_ok=True) if os.path.dirname(args.output) else None
    out_df.to_excel(args.output, index=False)
    print(f"Saved: {args.output}")

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo("Electoral Extractor", "Select input PDF file or a folder containing PDFs")
    input_path = filedialog.askopenfilename(title="Select a PDF (or Cancel to choose folder)", filetypes=[("PDF files","*.pdf")])
    if not input_path:
        input_path = filedialog.askdirectory(title="Select folder with PDFs")
        if not input_path:
            messagebox.showerror("Error", "No input selected."); return
    pdfs = gather_pdfs(input_path)
    if not pdfs:
        messagebox.showerror("Error", "No PDFs found in the selection."); return
    messagebox.showinfo("Output", "Choose where to save final_output.xlsx")
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="final_output.xlsx", filetypes=[("Excel","*.xlsx")])
    if not output_path:
        messagebox.showerror("Error", "No output selected."); return
    frames = []
    try:
        for p in pdfs:
            frames.append(parse_pdf(p))
        out_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        out_df.to_excel(output_path, index=False)
        messagebox.showinfo("Success", f"Saved: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def main():
    parser = argparse.ArgumentParser(description="Electoral Roll Data Extraction Tool (PDF ➜ Excel)")
    parser.add_argument("--input", type=str, help="Input PDF file or folder path")
    parser.add_argument("--output", type=str, help="Output Excel file path (e.g., C:\\out\\final_output.xlsx)")
    parser.add_argument("--gui", action="store_true", help="Launch GUI mode")
    args = parser.parse_args()
    if args.gui:
        run_gui(); return
    if not args.input or not args.output:
        parser.print_help(); sys.exit(1)
    run_cli(args)

if __name__ == "__main__":
    main()





