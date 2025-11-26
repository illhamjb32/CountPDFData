#!/usr/bin/env python3
"""
keyword_reports_by_company.py

Scans a target folder recursively for PDFs and counts keywords (category -> [keywords]).
Generates:
 - report_<MAIN_FOLDER>_<YYYY-MM-DD_HHMM>.csv    (detailed per file + keyword)
 - variable_report_<YYYY-MM-DD_HHMM>.csv         (Country,Company,Category,Keyword,2019..2024)
 - report_by_company_<YYYY-MM-DD_HHMM>.csv       (Country,Company,2019..2024) aggregated across keywords/categories
 - summary.txt                                    (number of unique PDF files scanned)
 - log_<...>.txt                                  (status per file)

Requires:
    pip install PyPDF2
Optional:
    pip install openpyxl
"""
from pathlib import Path
import argparse
import re
import csv
import sys
from datetime import datetime
from typing import Dict, List, Tuple
try:
    from PyPDF2 import PdfReader
except Exception:
    print("Missing dependency PyPDF2. Install with: python3 -m pip install PyPDF2")
    raise SystemExit(1)

# -------------------------
# Keywords (update as needed)
# -------------------------
keywords: Dict[str, List[str]] = {
    "Artificial Intelligence Technology": [
        "Artificial Intelligence", "Business Intelligence", "Image Understanding",
        "Investment Decision Support Systems", "Intelligent Data Analysis",
        "Intelligent Robots", "Machine Learning", "Deep Learning",
        "Semantic Search", "Biometric Technology", "Facial Recognition",
        "Speech Recognition", "Identity Verification", "Autonomous Driving",
        "Natural Language Processing"
    ],
    "Big data Technology": [
        "Big Data", "Data Mining", "Text Mining", "Data Visualization",
        "Heterogeneous Data", "Credit Reporting", "Augmented Reality",
        "Mixed Reality", "Virtual Reality"
    ],
    "Cloud Computing Technology": [
        "Cloud Computing", "Stream Computing", "Graph Computing",
        "Cyber-Physical Systems", "In-Memory Computing",
        "Multi-Party Secure Computing", "Neuromorphic Computing",
        "Green Computing", "Cognitive Computing", "Fusion Architecture",
        "Billion-Level Concurrency", "EB-Level Storage", "Internet of Things"
    ],
    "Block chain Technology / Digital Technology Applications": [
        "Block chain", "Digital Currency", "Differential Privacy Technology",
        "Smart Financial Contracts", "Mobile Internet", "Industrial Internet",
        "Mobile Interconnection", "Internet Healthcare", "E-commerce",
        "Mobile Payment", "Third-Party Payment", "NFC Payment",
        "Smart Energy", "B2B", "B2C", "C2C", "O2O", "Network Union",
        "Smart Wearables", "Smart Agriculture", "Smart Transportation",
        "Smart Healthcare", "Smart Customer Service", "Smart Home",
        "Robot-advisory", "Smart Tourism", "Smart Environmental Protection",
        "Smart Grid", "Smart Marketing", "Digital Marketing",
        "Unmanned Retail", "Internet Finance", "Digital Finance",
        "Fintech", "Financial Technology", "Quantitative Finance",
        "Open Banking"
    ],
    "Digital Technology Applications": [
        "Data Management", "Data Mining", "Data Networks", "Data Platforms",
        "Data Centers", "Data Science", "Digital Control", "Digital Technology",
        "Digital Communication", "Digital Networks", "Digital Intelligence",
        "Digital Terminals", "Digital Marketing", "Digitalization", "Big Data",
        "Cloud Computing", "Cloud IT", "Cloud Ecology", "Cloud Services",
        "Cloud Platforms", "Block chain", "Internet of Things", "Machine Learning"
    ],
    "Internet Business Models": [
        "Mobile Internet", "Industrial Internet", "Industry Internet",
        "Internet Solutions", "Internet Technology", "Internet Thinking",
        "Internet Actions", "Internet Business", "Internet Mobile",
        "Internet Applications", "Internet Marketing", "Internet Strategy",
        "Internet Platforms", "Internet Models", "Internet Business Models",
        "Internet Ecology", "E-commerce", "Electronic Commerce", "Online Offline",
        "Online to Offline", "O2O", "B2B", "C2C", "B2C"
    ],
    "Intelligent Manufacturing": [
        "Artificial Intelligence", "Advanced Intelligence", "Industrial Intelligence",
        "Mobile Intelligence", "Intelligent Control", "Intelligent Terminals",
        "Smart Mobility", "Intelligent Management", "Intelligent Factory",
        "Smart Logistics", "Intelligent Manufacturing", "Intelligent Warehousing",
        "Smart Technology", "Intelligent Devices", "Intelligent Production",
        "Intelligent Connected", "Intelligent Systems", "Intelligentization",
        "Automatic Control", "Automatic Monitoring", "Automatic inspection",
        "Automatic Detection", "Automatic Production", "Numerical Control",
        "Integration Standardization", "Integrated Solutions", "Integrated Control",
        "Integrated Systems", "Industrial Cloud", "Future Factory",
        "Lifecycle Management Manufacturing", "Execution Systems", "Virtualization",
        "Virtual Manufacturing"
    ],
    "Modern Information Systems": [
        "Information Sharing", "Information Management", "Information Integration",
        "Information Software", "Information Systems", "Information Networks",
        "Information Terminals", "Information Centers", "Informatization",
        "Networkization", "Industrial Information", "Industrial Communication"
    ],
}

# -------------------------
# Helpers
# -------------------------
def normalize_extracted_text(raw: str) -> str:
    """Normalize extracted text: remove hyphenation across lines, join lines, collapse whitespace, lowercase."""
    if not raw:
        return ""
    text = re.sub(r"-\s*\n\s*", "", raw)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

def extract_text_from_pdf(path_pdf: Path) -> Tuple[str, bool]:
    """Extract text from PDF using PyPDF2. Return (text, ok)."""
    parts: List[str] = []
    try:
        reader = PdfReader(str(path_pdf))
        for page in reader.pages:
            try:
                page_text = page.extract_text()
            except Exception:
                page_text = None
            if page_text:
                parts.append(page_text)
        return "\n".join(parts), True
    except Exception as e:
        return str(e), False

def build_patterns(keywords_data: Dict[str, List[str]], match_mode: str):
    patterns: List[Tuple[str, str, re.Pattern]] = []
    for cat, terms in keywords_data.items():
        for term in terms:
            t = term.strip()
            if not t:
                continue
            t_norm = t.lower()
            if match_mode == "substring":
                pat = re.compile(rf"(?=({re.escape(t_norm)}))", flags=re.IGNORECASE)
            else:
                pat = re.compile(rf"\b{re.escape(t_norm)}\b", flags=re.IGNORECASE)
            patterns.append((cat, t, pat))
    return patterns

def count_matches_in_text(text: str, pat: re.Pattern, match_mode: str) -> int:
    if not text:
        return 0
    if match_mode == "substring":
        matches = pat.findall(text)
        return len(matches)
    else:
        matches = pat.findall(text)
        return len(matches)

def find_year_from_filename(name: str) -> int:
    m = re.search(r"(19|20)\d{2}", name)
    if m:
        y = int(m.group(0))
        if 2019 <= y <= 2024:
            return y
    return None

def find_folder_by_name(root: Path, folder_name: str) -> Path:
    for p in root.rglob("*"):
        if p.is_dir() and p.name.lower() == folder_name.lower():
            return p
    raise FileNotFoundError(f"No folder named '{folder_name}' found under {root}")

def determine_main_folder(target_root: Path, pdf_parent: Path) -> str:
    try:
        rel = pdf_parent.relative_to(target_root)
        parts = rel.parts
        if len(parts) >= 1:
            return parts[0].split("(")[0].strip()
        else:
            return target_root.name.split("(")[0].strip()
    except Exception:
        return target_root.name.split("(")[0].strip()

def determine_company_name(pdf_parent: Path) -> str:
    # per your request: company name should be truncated before '('
    raw = pdf_parent.name or "."
    return raw.split("(")[0].strip()

# -------------------------
# Scanning & aggregation
# -------------------------
def scan_folder_for_pdfs(base_folder: Path, patterns: List[Tuple[str,str,re.Pattern]], match_mode: str) -> Tuple[List[List], List[str]]:
    """
    Recursively scan base_folder for PDFs, return:
      - rows: list of [main_folder, sub_folder, filename, category, keyword, count]
      - log_lines: list of log strings
    """
    rows: List[List] = []
    log_lines: List[str] = []
    scanned_files_set = set()

    pdf_files = sorted(base_folder.rglob("*.pdf"))
    if not pdf_files:
        return rows, log_lines

    for pdf in pdf_files:
        raw, ok = extract_text_from_pdf(pdf)
        sub_folder_raw = pdf.parent.name or "."
        sub_folder = determine_company_name(pdf.parent)  # cleaned company name
        main_folder = determine_main_folder(base_folder, pdf.parent)
        scanned_files_set.add(pdf.name)

        if not ok:
            # if extraction failed, mark zeros and log error
            total_found_for_file = 0
            for cat, term_original, pat in patterns:
                rows.append([main_folder, sub_folder, pdf.name, cat, term_original, 0])
            log_lines.append(f"{main_folder}_{sub_folder}_{pdf.name}_status : Error ; {total_found_for_file}")
            print(f"Scanned: {pdf.name} — error")
            continue

        # successful extraction
        normalized = normalize_extracted_text(raw)
        total_found_for_file = 0
        for cat, term_original, pat in patterns:
            cnt = count_matches_in_text(normalized, pat, match_mode)
            rows.append([main_folder, sub_folder, pdf.name, cat, term_original, cnt])
            if cnt > 0:
                total_found_for_file += 1

        log_lines.append(f"{main_folder}_{sub_folder}_{pdf.name}_status : Done ; {total_found_for_file}")
        print(f"Scanned: {pdf.name} — done")

    return rows, log_lines

# -------------------------
# Save CSVs & reports
# -------------------------
def save_csv(path: Path, rows: List[List]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(["main_folder", "sub_folder", "filename", "category", "keyword", "count"])
        writer.writerows(rows)

def write_summary(path: Path, files_scanned_count: int) -> None:
    summary_path = path.parent / "summary.txt"
    with open(summary_path, "w", encoding="utf-8") as s:
        s.write(str(files_scanned_count))

def write_log(path: Path, log_lines: List[str]) -> None:
    log_name = f"log_{path.parent.name}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.txt"
    log_path = path.parent / log_name
    with open(log_path, "w", encoding="utf-8") as L:
        for line in log_lines:
            L.write(line + "\n")

def try_save_xlsx(path: Path, rows: List[List]) -> None:
    try:
        from openpyxl import Workbook
    except Exception:
        return
    wb = Workbook()
    ws = wb.active
    ws.append(["main_folder", "sub_folder", "filename", "category", "keyword", "count"])
    for r in rows:
        ws.append(r)
    wb.save(str(path.with_suffix(".xlsx")))

# -------------------------
# Generate variable_report (detailed per category+keyword per year)
# -------------------------
def generate_variable_report(rows: List[List], target_folder: Path) -> Path:
    """
    Create variable_report CSV with:
    Country;Company Name;Category;Keyword;2019-2024
    """
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H%M")
    out_name = f"variable_report_{timestamp}.csv"
    out_path = target_folder / out_name

    agg: Dict[Tuple[str,str,str,str], Dict[int,int]] = {}

    for main_folder, sub_folder, filename, cat, keyword, count in rows:
        m = re.search(r"(19|20)\d{2}", filename)
        if not m:
            continue
        year = int(m.group(0))
        if not (2019 <= year <= 2024):
            continue
        key = (main_folder, sub_folder, cat, keyword)
        if key not in agg:
            agg[key] = {y:0 for y in range(2019,2025)}
        agg[key][year] += int(count)

    # write CSV
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(["Country","Company Name","Category","Keyword","2019","2020","2021","2022","2023","2024"])
        for (country, company, cat, keyword), year_map in sorted(agg.items(), key=lambda k: (k[0][0].lower(), k[0][1].lower())):
            row = [country, company, cat, keyword] + [year_map[y] for y in range(2019,2025)]
            w.writerow(row)

    return out_path

# -------------------------
# Generate report_by_company (aggregate across keywords & categories)
# -------------------------
def generate_report_by_company(rows: List[List], target_folder: Path) -> Path:
    """
    Create report_by_company CSV with:
    Country;Company Name;2019;2020;2021;2022;2023;2024
    Aggregation sums counts across all categories & keywords for that company.
    """
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H%M")
    out_name = f"report_by_company_{timestamp}.csv"
    out_path = target_folder / out_name

    agg: Dict[Tuple[str,str], Dict[int,int]] = {}  # (country, company) -> year->count

    for main_folder, sub_folder, filename, cat, keyword, count in rows:
        m = re.search(r"(19|20)\d{2}", filename)
        if not m:
            continue
        year = int(m.group(0))
        if not (2019 <= year <= 2024):
            continue
        key = (main_folder, sub_folder)
        if key not in agg:
            agg[key] = {y:0 for y in range(2019,2025)}
        agg[key][year] += int(count)

    # write CSV
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(["Country","Company Name","2019","2020","2021","2022","2023","2024"])
        for (country, company), year_map in sorted(agg.items(), key=lambda k: (k[0][0].lower(), k[0][1].lower())):
            row = [country, company] + [year_map[y] for y in range(2019,2025)]
            w.writerow(row)

    return out_path

# -------------------------
# CLI & Main
# -------------------------
def parse_args():
    p = argparse.ArgumentParser(description="Keyword scanner + variable_report + report_by_company")
    p.add_argument("-d", "--dir", default=".", help="Folder path to scan (default: current folder). Scans recursively.")
    p.add_argument("-r", "--root", default=None, help="Root to search for named folder (use with -n).")
    p.add_argument("-n", "--name", default=None, help="If provided, find folder by this name under --root and scan it.")
    p.add_argument("--xlsx", action="store_true", help="Also save an .xlsx (requires openpyxl).")
    p.add_argument("--match", choices=["whole","substring"], default="whole", help="Matching mode")
    return p.parse_args()

def main():
    args = parse_args()

    if args.name and args.root:
        root = Path(args.root).expanduser().resolve()
        try:
            target_folder = find_folder_by_name(root, args.name)
        except FileNotFoundError as e:
            print(f"[ERROR] {e}", file=sys.stderr)
            return
    else:
        target_folder = Path(args.dir).expanduser().resolve()

    if not target_folder.exists() or not target_folder.is_dir():
        print(f"[ERROR] Target folder does not exist or is not a directory: {target_folder}", file=sys.stderr)
        return

    patterns = build_patterns(keywords, args.match)
    rows, log_lines = scan_folder_for_pdfs(target_folder, patterns, args.match)

    # report_...csv (detailed per file + keyword)
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H%M")
    main_for_filename = target_folder.name
    out_csv = target_folder / f"report_{main_for_filename}_{timestamp}.csv"
    save_csv(out_csv, rows)

    # summary & log
    unique_pdfs = len(set(r[2] for r in rows))
    write_summary(out_csv, unique_pdfs)
    write_log(out_csv, log_lines)

    if args.xlsx:
        try_save_xlsx(out_csv, rows)

    # minimal final summary + generate additional reports
    nonzero = sum(1 for r in rows if r[5] > 0)

    var_path = generate_variable_report(rows, target_folder)
    byco_path = generate_report_by_company(rows, target_folder)

    print(f"Scan complete. Files scanned: {unique_pdfs}. Non-zero matches: {nonzero}.")
    print(f"Report saved to: {out_csv}")
    print(f"Variable report saved to: {var_path}")
    print(f"Report by company saved to: {byco_path}")
    print("Summary.txt & log file generated.")

if __name__ == "__main__":
    main()
