#!/usr/bin/env python3
"""
rescaner.py — v2 (optimized, multiprocess)

Scans a target folder recursively for PDFs, counts keyword frequencies AND
total words per file, then generates frequency + ratio reports.

Outputs (written into the target folder):
 - report_<FOLDER>_<ts>.csv      frequency per file x keyword (+ total_words)
 - ratio_report_<ts>.csv         ratio = keyword count / total words, per file x keyword
 - variable_report_<ts>.csv      frequency: Country x Company x Category x Keyword x 2019..2024
 - report_by_company_<ts>.xlsx   3 sheets: frekuensi_keyword / sum_word / rasio (Country x Company x 2019..2025)
 - summary.txt                   number of unique PDFs scanned
 - log_<...>.txt                 status per file (+ reason on error, + no-year warnings)

Requires:  pip install pypdf
Optional fallbacks (auto-detected if installed): PyMuPDF, pdfminer.six, pikepdf

Usage:
    python rescaner.py -d "C:/path/to/reports"
    python rescaner.py -r "C:/data" -n "China" --match substring --workers 8
"""
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
import argparse
import csv
import json
import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Tuple

try:
    from pypdf import PdfReader
except Exception:
    print("Missing dependency pypdf. Install with: python -m pip install pypdf")
    raise SystemExit(1)

YEARS = list(range(2019, 2026))  # 2019..2025
YEAR_RE = re.compile(r"(?:19|20)\d{2}")

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
    """Remove hyphenation across line breaks, join lines, collapse whitespace, lowercase."""
    if not raw:
        return ""
    text = re.sub(r"-\s*\n\s*", "", raw)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

def _ocr_pdf(path_pdf: Path) -> str:
    """OCR fallback for image-based (scanned) PDFs: render pages via PyMuPDF, read via RapidOCR.
    Pages are OCR'd in parallel via a thread pool (OCR releases the GIL in onnxruntime)."""
    import pymupdf
    from rapidocr_onnxruntime import RapidOCR
    from concurrent.futures import ThreadPoolExecutor
    doc = pymupdf.open(str(path_pdf))
    n = len(doc)

    def render(i: int) -> str:
        ocr = RapidOCR()
        pg = doc[i]
        pix = pg.get_pixmap(dpi=200)
        img_path = str(path_pdf) + f"_ocr_p{i}.png"
        pix.save(img_path)
        try:
            res, _ = ocr(img_path)
            return " ".join(r[1] for r in res) if res else ""
        finally:
            try:
                os.remove(img_path)
            except OSError:
                pass

    threads = min(8, (os.cpu_count() or 4))
    with ThreadPoolExecutor(max_workers=threads) as ex:
        parts = list(ex.map(render, range(n)))
    return "\n".join(parts)

def extract_text_from_pdf(path_pdf: Path) -> Tuple[str, bool]:
    """Extract text: PyMuPDF (fastest) -> pypdf -> pdfminer -> pikepdf rewrite -> OCR (image PDFs)."""
    err = "unknown error"
    # 1) PyMuPDF — fastest by far for big annual reports
    try:
        import pymupdf
        doc = pymupdf.open(str(path_pdf))
        t = "\n".join(pg.get_text() for pg in doc)
        if t.strip():
            return t, True
        err = "pymupdf: empty text"
    except Exception as e:
        err = f"pymupdf: {e}"

    # 2) pypdf
    try:
        reader = PdfReader(str(path_pdf))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                pass
        text = "\n".join((p.extract_text() or "") for p in reader.pages)
        if text.strip():
            return text, True
    except Exception:
        pass

    # 3) pdfminer.six
    try:
        from pdfminer.high_level import extract_text as pdfminer_extract
        t = pdfminer_extract(str(path_pdf))
        if t.strip():
            return t, True
    except Exception:
        pass

    # 4) pikepdf: rewrite (repairs encryption/xref damage), then retry extractors
    tmp = None
    try:
        import pikepdf
        tmp = str(path_pdf) + "_dec.pdf"
        with pikepdf.open(str(path_pdf)) as pdf:
            pdf.save(tmp)
        try:
            import pymupdf
            doc = pymupdf.open(tmp)
            t = "\n".join(pg.get_text() for pg in doc)
            if t.strip():
                return t, True
        except Exception:
            pass
        try:
            from pdfminer.high_level import extract_text as pdfminer_extract
            t = pdfminer_extract(tmp)
            if t.strip():
                return t, True
        except Exception:
            pass
    except Exception:
        pass
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass

    # 5) OCR — image-based/scanned PDF (all text extractors returned empty)
    try:
        t = _ocr_pdf(path_pdf)
        if t.strip():
            return t + "\n[[OCR]]", True  # marker: downstream flags this file as OCR-sourced
        err = "ocr: empty text (fully image-based?)"
    except Exception as e:
        err = f"ocr: {e}"

    return err, False

def build_pattern(match_mode: str) -> Tuple[List[Tuple[str, str]], re.Pattern]:
    """
    ONE combined regex for all keywords (single pass over the text instead of
    ~130 separate scans). Terms are sorted longest-first so e.g. "Internet Banking"
    is counted as itself and never double-counted again as "Internet".
    Each term is its own capture group; m.lastindex identifies which one hit.
    """
    terms = [(cat, kw.strip()) for cat, kws in keywords.items() for kw in kws if kw.strip()]
    terms.sort(key=lambda c: len(c[1]), reverse=True)
    if match_mode == "substring":
        pat = re.compile("|".join(rf"(?=({re.escape(kw.lower())}))" for _, kw in terms),
                        flags=re.IGNORECASE)
    else:
        pat = re.compile("|".join(rf"(\b{re.escape(kw.lower())}\b)" for _, kw in terms),
                        flags=re.IGNORECASE)
    return terms, pat

def find_year(filename: str):
    """First year-token in 2019..2024 found anywhere in the filename (None if absent)."""
    for tok in YEAR_RE.findall(filename):
        y = int(tok)
        if y in YEARS:
            return y
    return None

def find_folder_by_name(root: Path, folder_name: str) -> Path:
    for p in root.rglob("*"):
        if p.is_dir() and p.name.lower() == folder_name.lower():
            return p
    raise FileNotFoundError(f"No folder named '{folder_name}' found under {root}")

def determine_main_folder(target_root: Path, pdf_parent: Path) -> str:
    """Country label = the scanned root folder name, cleaned of (...) suffixes."""
    return target_root.name.split("(")[0].strip()

def determine_company_name(pdf_parent: Path) -> str:
    """Company label = immediate parent folder name, cleaned of (...) suffixes."""
    raw = pdf_parent.name or "."
    return raw.split("(")[0].strip()

def determine_scan_root(pdf_path: Path, fallback: Path) -> Path:
    """Country root = the folder that directly contains the country folder.

    Layout: <Annual Report>/<Country>/<Company>/<file.pdf>. The country folder is
    the PDF's grandparent and the company folder its parent. We verify the
    grandparent really contains the company folder (i.e. we are inside a country),
    otherwise (PDF directly under the root, or deeper nesting) fall back to
    walking up until we leave the scanner's own folder.
    """
    script_folder = Path(__file__).resolve().parent
    parent = pdf_path.parent
    grand = parent.parent
    if grand != script_folder and grand.is_dir() and parent.name != grand.name:
        return grand
    for p in pdf_path.parents:
        if p != script_folder:
            return p
    return fallback

# -------------------------
# Worker (runs in its own process)
# -------------------------
_CTX: Dict = {}
_CKPT: Dict = {}  # absolute path -> result dict (loaded from checkpoint file at startup)

def _init_worker(match_mode: str, ckpt: Dict = None) -> None:
    terms, pat = build_pattern(match_mode)
    _CTX["terms"] = terms
    _CTX["pat"] = pat
    if ckpt:
        _CKPT.update(ckpt)

def _process_pdf(pdf_str: str, base_str: str) -> Dict:
    pdf = Path(pdf_str)
    base = determine_scan_root(pdf, Path(base_str))
    main_folder = determine_main_folder(base, pdf.parent)
    sub_folder = determine_company_name(pdf.parent)

    # resume: skip files already scanned in a previous run (checkpoint)
    ckpt_key = str(pdf)
    if ckpt_key in _CKPT:
        return _CKPT[ckpt_key]

    raw, ok = extract_text_from_pdf(pdf)
    if not ok:
        return {"file": pdf.name, "main": main_folder, "sub": sub_folder,
                "ok": False, "error": " ".join(str(raw).split())}
    is_ocr = "[[OCR]]" in raw
    if is_ocr:
        raw = raw.replace("[[OCR]]", "")
        # OCR noise filter: drop tokens that are >60% non-alphanumeric garbage
        raw = " ".join(w for w in raw.split()
                       if sum(c.isalnum() for c in w) >= len(w) * 0.4 or len(w) <= 2)

    text = normalize_extracted_text(raw)
    total_words = len(text.split())

    counts = [0] * len(_CTX["terms"])
    for m in _CTX["pat"].finditer(text):
        gi = m.lastindex
        if gi:
            counts[gi - 1] += 1

    rows = [[main_folder, sub_folder, pdf.name, cat, kw, c, total_words]
            for (cat, kw), c in zip(_CTX["terms"], counts)]
    return {"file": pdf.name, "main": main_folder, "sub": sub_folder,
            "ok": True, "rows": rows, "total_words": total_words,
            "total_hits": sum(counts), "ocr": is_ocr}

# -------------------------
# Reports
# -------------------------
def save_csv(path: Path, header: List[str], rows: List[List]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(header)
        w.writerows(rows)

def write_summary(out_dir: Path, files_scanned_count: int) -> None:
    with open(out_dir / "summary.txt", "w", encoding="utf-8") as s:
        s.write(str(files_scanned_count))

def write_log(out_dir: Path, log_lines: List[str], ts: str) -> None:
    with open(out_dir / f"log_report_{ts}.txt", "w", encoding="utf-8") as L:
        L.write(f"=== Run {datetime.now().strftime('%Y-%m-%d %H:%M')} ===\n")
        for line in log_lines:
            L.write(line + "\n")

def generate_variable_report(rows: List[List], target_folder: Path, ts: str) -> Path:
    """Frequency: Country;Company;Category;Keyword;2019..2024"""
    out_path = target_folder / f"variable_report_{ts}.csv"
    agg: Dict[Tuple[str, str, str, str], Dict[int, int]] = {}
    for main_folder, sub_folder, filename, cat, kw, count, _tw in rows:
        y = find_year(filename)
        if y is None:
            continue
        key = (main_folder, sub_folder, cat, kw)
        agg.setdefault(key, {yy: 0 for yy in YEARS})[y] += count
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Country", "Company Name", "Category", "Keyword"] + [str(y) for y in YEARS])
        for (country, company, cat, kw), ym in sorted(agg.items()):
            w.writerow([country, company, cat, kw] + [ym[y] for y in YEARS])
    return out_path

def generate_report_by_company(rows: List[List], target_folder: Path, ts: str) -> Path:
    """3-sheet .xlsx: frekuensi_keyword, sum_word, rasio (freq/words, 3 decimals)."""
    out_path = target_folder / f"report_by_company_{ts}.xlsx"
    try:
        from openpyxl import Workbook
    except Exception:
        print("[WARN] openpyxl not installed; skipping report_by_company. "
              "Install with: python -m pip install openpyxl", file=sys.stderr)
        return out_path

    freq: Dict[Tuple[str, str], Dict[int, int]] = {}
    words: Dict[Tuple[str, str], Dict[int, int]] = {}
    seen_files: set = set()
    for main_folder, sub_folder, filename, _cat, _kw, count, tw in rows:
        y = find_year(filename)
        if y is None:
            continue
        key = (main_folder, sub_folder)
        freq.setdefault(key, {yy: 0 for yy in YEARS})[y] += count
        fkey = (main_folder, sub_folder, filename)  # count each file's words once
        if fkey not in seen_files:
            seen_files.add(fkey)
            words.setdefault(key, {yy: 0 for yy in YEARS})[y] += tw

    header = ["Country", "Company Name"] + [str(y) for y in YEARS]
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "frekuensi_keyword"
    ws1.append(header)
    for key in sorted(freq):
        ws1.append([key[0], key[1]] + [freq[key][y] for y in YEARS])

    ws2 = wb.create_sheet("sum_word")
    ws2.append(header)
    for key in sorted(words):
        ws2.append([key[0], key[1]] + [words[key][y] for y in YEARS])

    ws3 = wb.create_sheet("rasio")
    ws3.append(header)
    for key in sorted(freq):
        row = [key[0], key[1]]
        for y in YEARS:
            w = words.get(key, {}).get(y, 0)
            row.append((freq[key][y] / w) if w else 0.0)
        ws3.append(row)
        for cell in ws3[ws3.max_row][2:]:
            cell.number_format = "0.00000"
    wb.save(str(out_path))
    return out_path

def parse_args():
    p = argparse.ArgumentParser(description="Keyword scanner v2: frequency + ratio reports, multiprocess")
    p.add_argument("-d", "--dir", default=".", help="Folder path to scan (default: current folder). Scans recursively.")
    p.add_argument("-r", "--root", default=None, help="Root to search for named folder (use with -n).")
    p.add_argument("-n", "--name", default=None, help="If provided, find folder by this name under --root and scan it.")
    p.add_argument("--match", choices=["whole", "substring"], default="whole", help="Matching mode (default: whole word).")
    p.add_argument("-w", "--workers", type=int, default=0, help="Number of worker processes (0 = auto = CPU count).")
    p.add_argument("--fresh", action="store_true", help="Ignore checkpoint and rescan everything from scratch.")
    return p.parse_args()

CKPT_FILE = Path(__file__).resolve().parent / "reports" / "checkpoint.jsonl"

def load_checkpoint() -> Dict:
    """Resume support: {abs_path: result} from the previous (interrupted) run."""
    ckpt: Dict = {}
    if CKPT_FILE.exists():
        with open(CKPT_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                    ckpt[rec["path"]] = rec["result"]
                except (json.JSONDecodeError, KeyError):
                    continue  # torn last line from a crash — safe to ignore
    return ckpt

def append_checkpoint(path_str: str, result: Dict) -> None:
    """One JSON line per finished file, appended immediately (crash-safe)."""
    CKPT_FILE.parent.mkdir(exist_ok=True)
    with open(CKPT_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps({"path": path_str, "result": result}, ensure_ascii=False) + "\n")

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

    pdf_files = sorted(target_folder.rglob("*.pdf"))
    if not pdf_files:
        print("[ERROR] No PDF files found in target folder.", file=sys.stderr)
        return

    ckpt = {} if args.fresh else load_checkpoint()
    todo = [p for p in pdf_files if str(p) not in ckpt]
    if ckpt and todo:
        print(f"Resume: {len(pdf_files) - len(todo)} file(s) already scanned, {len(todo)} left.")
    elif ckpt and not todo:
        print("All files already scanned (checkpoint complete). Regenerating reports...")
    if args.fresh and CKPT_FILE.exists():
        CKPT_FILE.unlink()

    workers = args.workers if args.workers > 0 else (os.cpu_count() or 4)
    print(f"Scanning {len(todo)} PDFs with {workers} workers (match={args.match})...")

    results: List[Dict] = [ckpt[str(p)] for p in pdf_files if str(p) in ckpt]
    all_results: Dict[str, Dict] = dict(ckpt)  # every file ever scanned (all countries)
    t0 = datetime.now()
    if todo:
        with ProcessPoolExecutor(max_workers=workers, initializer=_init_worker,
                                 initargs=(args.match, ckpt)) as ex:
            futs = {ex.submit(_process_pdf, str(p), str(target_folder)): p for p in todo}
            done = 0
            for fut in as_completed(futs):
                r = fut.result()
                results.append(r)
                all_results[str(futs[fut])] = r
                append_checkpoint(str(futs[fut]), r)
                done += 1
                elapsed = (datetime.now() - t0).total_seconds()
                eta = elapsed / done * (len(todo) - done)
                status = "done" if r["ok"] else f"error ({r['error']})"
                ocr_tag = " [OCR]" if r.get("ocr") else ""
                print(f"[{done}/{len(todo)}] {r['file']}{ocr_tag} — {status} "
                      f"| elapsed {int(elapsed//60)}m{int(elapsed%60):02d}s | ETA {int(eta//60)}m{int(eta%60):02d}s")

    results.sort(key=lambda r: r["file"])  # deterministic output order
    rows: List[List] = []
    log_lines: List[str] = []
    no_year_files: List[str] = []
    for r in results:
        if not r["ok"]:
            rows.extend([r["main"], r["sub"], r["file"], cat, kw, 0, 0]
                        for cat, kw in [(c, k) for c, kws in keywords.items() for k in kws])
            log_lines.append(f"{r['main']}_{r['sub']}_{r['file']}_status : Error ; 0 ; words=0 ; Reason: {r['error']}")
        else:
            rows.extend(r["rows"])
            ocr_tag = " ; ocr=true" if r.get("ocr") else ""
            log_lines.append(f"{r['main']}_{r['sub']}_{r['file']}_status : Done ; {r['total_hits']} ; words={r['total_words']}{ocr_tag}")
            if find_year(r["file"]) is None:
                no_year_files.append(r["file"])
    for f in no_year_files:
        log_lines.append(f"WARNING: '{f}' has no year 2019-2025 in filename; excluded from year aggregations.")

    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    # Reports always land in the scanner's own folder (ScanData/reports), NOT inside
    # the scanned tree — keeps the Annual Report folders clean.
    out_dir = Path(__file__).resolve().parent / "reports"
    out_dir.mkdir(exist_ok=True)
    unique_pdfs = len(pdf_files)  # count actual files, not distinct filenames
    write_summary(out_dir, unique_pdfs)
    write_log(out_dir, log_lines, ts)

    # One file per report category, accumulated over ALL scanned countries.
    # Fixed filenames: each run overwrites the previous, so there is always
    # exactly one variable_report.csv and one report_by_company.xlsx.
    rows_all: List[List] = []
    for r in sorted(all_results.values(), key=lambda x: (x["main"], x["sub"], x["file"])):
        if r["ok"]:
            rows_all.extend(r["rows"])
        else:
            rows_all.extend([r["main"], r["sub"], r["file"], cat, kw, 0, 0]
                            for cat, kws in keywords.items() for kw in kws)

    var_path = generate_variable_report(rows_all, out_dir, ts)
    byco_path = generate_report_by_company(rows_all, out_dir, ts)

    nonzero = sum(1 for r in rows if r[5] > 0)
    print(f"\nScan complete. Files scanned: {unique_pdfs} ({len(all_results)} total in reports). Non-zero keyword rows: {nonzero}.")
    print(f"Variable report  : {var_path}")
    print(f"Report by company: {byco_path}")

if __name__ == "__main__":
    main()
