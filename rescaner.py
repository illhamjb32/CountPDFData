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
    "Information Technology": [
        "Informatization Construction", "Information Science", "IT Governance",
        "IT Architecture", "Online", "Quick Response Code", "Opening",
        "Informatization", "Automation", "Digitalization", "Intelligent",
        "Scene", "Instant Messaging", "5G", "Information System",
        "Information Security", "Open", "Interconnect", "Sharing",
        "Virtual Reality", "Cyber-Physical Systems", "FinTech",
        "Financial Technology"
    ],
    "Artificial Intelligence": [
        "Artificial Intelligence", "Face Recognition", "Real-time Monitoring",
        "Fingerprint Identification", "Deep Learning", "Wearable",
        "Intelligent", "Smart", "Machine Learning", "Face Swiping",
        "Voiceprint", "Intelligent Speech", "Biometric Identification",
        "Biometric Authentication", "Text Mining", "Brain-inspired Computing",
        "Image Understanding", "Natural Language Processing"
    ],
    "Blockchain Technology": [
        "Blockchain", "Alliance Chain", "Secure Multi-Party Computation",
        "Distributed Computation"
    ],
    "Cloud Technology": [
        "Cloud Computing", "Cloud Serving", "Finance Cloud",
        "Cloud Computing Architecture", "IaaS", "PaaS", "SaaS"
    ],
    "Data Technology": [
        "Big Data", "Data Mining", "Data Stream", "Data Set",
        "Information Mining", "CRM"
    ],
    "Internet Technology": [
        "Internet", "Cellphone", "Mobile", "Mobile Device", "Network",
        "Remote", "Electronic", "API", "Internet of Things",
        "Mobile Communications", "Internet Finance", "Biosphere",
        "Open Banking", "Online Banking", "Electronic Banking", "E-Banking",
        "Internet Banking", "Mobile Banking", "Electronic Wallet",
        "WeChat Banking", "Self-service Equipment", "E-finance",
        "Smart Banking", "Online Financial Products", "VTM",
        "Electronic Commerce", "E-commerce", "Open Internet Banking",
        "Open System Interconnection", "Digital Banking",
        "Online Supply Chain", "Intelligent Retail", "Contactless Commerce",
        "Scene Finance", "Third Party Payment", "Mobile Payment",
        "Online Payment", "Net Payment", "Mobile Phone and Payment",
        "NFC Payment", "Digital Currency"
    ],
    "Business (Service) Channels": [
        "Internet Finance", "Biosphere", "Open Banking", "Online Banking",
        "Electronic Banking", "E-banking", "Internet Banking", "Mobile Banking",
        "Electronic Wallet", "WeChat Banking", "Self-service Equipment",
        "E-finance", "Smart Banking", "Online Financial Products", "VTM",
        "Electronic Commerce", "E-commerce", "Open Internet Banking",
        "Open System Interconnection", "Digital Banking", "Online Supply Chain",
        "Intelligent Retail", "Contactless Commerce", "Scene Finance",
        "Third Party Payment", "Mobile Payment", "Online Payment",
        "Net Payment", "Mobile Phone and Payment", "NFC Payment",
        "Digital Currency", "Electronic Payment", "Barcode Payment",
        "Two-dimensional Barcode Payment"
    ],
    "Gross Settlement": [
        "Electronic Payment", "Barcode Payment", "Two-dimensional Barcode Payment",
        "EB-class Storage", "Wearable Payment", "Senseless Payment"
    ],
    "Resource Allocation": [
        "Internet Financing", "Peer-to-peer Lending", "P2P Lending",
        "Crowdfunding", "Internet Lending", "Network Financing",
        "Online Investment", "Equity-based Crowdfunding", "Investment Decision Aid System",
        "Online Financing", "Financial Inclusion", "Personalized Pricing",
        "Scene Financing"
    ],
    "Financial Management": [
        "Consumer Finance", "Online Wealth Management", "Online Insurance",
        "Robot Financing", "Expert Advisor", "Intelligent Advisor"
    ],
    "Risk Management": [
        "Big Data Credit", "Big Data Risk Control"
    ]
}


# -------------------------
# Helpers
# -------------------------
def normalize_extracted_text(raw: str) -> str:
    if not raw:
        return ""
    text = re.sub(r"-\s*\n\s*", "", raw)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

def extract_text_from_pdf(path_pdf: Path) -> Tuple[str, bool]:
    # 1️⃣ Try PyPDF2 first (preserve original behavior)
    error_pypdf2 = "Unknown PyPDF2 error"
    try:
        reader = PdfReader(str(path_pdf))
        parts: List[str] = []
        for page in reader.pages:
            try:
                page_text = page.extract_text()
            except Exception:
                page_text = None
            if page_text:
                parts.append(page_text)
        if parts:
            return "\n".join(parts), True
        error_pypdf2 = "PyPDF2 returned empty text"
    except Exception as e:
        error_pypdf2 = str(e)

    # 2️⃣ Try fallback extractors (import locally to avoid top-level import errors)
    try:
        from pdfminer.high_level import extract_text as pdfminer_extract
    except Exception:
        pdfminer_extract = None

    try:
        import fitz  # PyMuPDF
    except Exception:
        fitz = None

    try:
        import pikepdf
    except Exception:
        pikepdf = None

    # helper for pdfminer
    if pdfminer_extract:
        try:
            text = pdfminer_extract(str(path_pdf))
            if text and text.strip():
                return text, True
        except Exception:
            pass

    # helper for pymupdf
    if fitz:
        try:
            doc = fitz.open(str(path_pdf))
            t = "".join(page.get_text() for page in doc)
            if t and t.strip():
                return t, True
        except Exception:
            pass

    # try decrypt + pdfminer / pymupdf if pikepdf available
    if pikepdf:
        try:
            tmp = str(path_pdf) + "_decrypted.pdf"
            with pikepdf.open(str(path_pdf)) as pdf:
                pdf.save(tmp)
            # try pdfminer on decrypted
            if pdfminer_extract:
                try:
                    text = pdfminer_extract(tmp)
                    if text and text.strip():
                        return text, True
                except Exception:
                    pass
            # try pymupdf on decrypted
            if fitz:
                try:
                    doc = fitz.open(tmp)
                    t = "".join(page.get_text() for page in doc)
                    if t and t.strip():
                        return t, True
                except Exception:
                    pass
        except Exception:
            pass

    # still failed
    return error_pypdf2, False

def build_patterns(keywords_data: Dict[str, List[str]], match_mode: str):
    patterns: List[Tuple[str, str, re.Pattern]] = []
    for cat, terms in keywords_data.items():
        for term in terms:
            t = term.strip()
            if not t:
                continue
            if match_mode == "substring":
                pat = re.compile(rf"(?=({re.escape(t.lower())}))", flags=re.IGNORECASE)
            else:
                pat = re.compile(rf"\b{re.escape(t.lower())}\b", flags=re.IGNORECASE)
            patterns.append((cat, t, pat))
    return patterns

def count_matches_in_text(text: str, pat: re.Pattern, match_mode: str) -> int:
    if not text:
        return 0
    matches = pat.findall(text)
    return len(matches)

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
    raw = pdf_parent.name or "."
    return raw.split("(")[0].strip()


# -------------------------
# Scanning & aggregation
# -------------------------
def scan_folder_for_pdfs(base_folder: Path, patterns: List[Tuple[str,str,re.Pattern]], match_mode: str):
    rows: List[List] = []
    log_lines: List[str] = []
    scanned_files_set = set()

    pdf_files = sorted(base_folder.rglob("*.pdf"))
    if not pdf_files:
        return rows, log_lines

    for pdf in pdf_files:
        raw, ok = extract_text_from_pdf(pdf)
        sub_folder = determine_company_name(pdf.parent)
        main_folder = determine_main_folder(base_folder, pdf.parent)
        scanned_files_set.add(pdf.name)

        if not ok:
            error_msg = raw.strip().replace("\n", " ")
            total_found_for_file = 0

            for cat, term_original, pat in patterns:
                rows.append([main_folder, sub_folder, pdf.name, cat, term_original, 0])

            log_lines.append(
                f"{main_folder}_{sub_folder}_{pdf.name}_status : Error ; {total_found_for_file} ; Reason: {error_msg}"
            )
            print(f"Scanned: {pdf.name} — error ({error_msg})")
            continue

        normalized = normalize_extracted_text(raw)
        total_found_for_file = 0

        for cat, term_original, pat in patterns:
            cnt = count_matches_in_text(normalized, pat, match_mode)
            rows.append([main_folder, sub_folder, pdf.name, cat, term_original, cnt])
            if cnt > 0:
                total_found_for_file += cnt

        log_lines.append(f"{main_folder}_{sub_folder}_{pdf.name}_status : Done ; {total_found_for_file}")
        print(f"Scanned: {pdf.name} — done")

    return rows, log_lines


# -------------------------
# Save & report
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

def generate_variable_report(rows: List[List], target_folder: Path) -> Path:
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

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(["Country","Company Name","Category","Keyword","2019","2020","2021","2022","2023","2024"])
        for (country, company, cat, keyword), year_map in sorted(agg.items()):
            row = [country, company, cat, keyword] + [year_map[y] for y in range(2019,2025)]
            w.writerow(row)

    return out_path

def generate_report_by_company(rows: List[List], target_folder: Path) -> Path:
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H%M")
    out_name = f"report_by_company_{timestamp}.csv"
    out_path = target_folder / out_name

    agg: Dict[Tuple[str,str], Dict[int,int]] = {}

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

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(["Country","Company Name","2019","2020","2021","2022","2023","2024"])
        for (country, company), year_map in sorted(agg.items()):
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
