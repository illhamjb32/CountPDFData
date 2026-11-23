#!/usr/bin/env python3
# Download file yang kurang dari Google Drive -> lokal. Resumable: skip file yang sudah ada.
import re, subprocess, html, json, os, sys, time

ROOT_ID = "1un48qFqzXgkY7mBtn2WYTp6VIyNteRyZ"
LOCAL_ROOT = "/Users/ilhamjb/Documents/Project/Python/CountData/Annual Report"
CACHE = "/tmp/drive_tree_ids.json"
LOG = open("/Users/ilhamjb/Documents/Project/Python/CountData/download_log.txt", "a")

def log(msg):
    print(msg, flush=True); LOG.write(msg + "\n"); LOG.flush()

def curl(url):
    return subprocess.run(["curl", "-sL", "--max-time", "60", url],
                          capture_output=True, text=True).stdout

def parse_folder(fid):
    h = curl(f"https://drive.google.com/embeddedfolderview?id={fid}#list")
    out = []
    for m in re.finditer(r'<div class="flip-entry" id="entry-([A-Za-z0-9_-]+)".*?<div class="flip-entry-title">([^<]*)</div>', h, re.S):
        out.append((m.group(1), html.unescape(m.group(2)).strip()))
    return out

def build_tree():
    if os.path.exists(CACHE):
        return json.load(open(CACHE))
    tree = {}  # country -> bank -> [(file_id, filename)]
    for cid, cname in parse_folder(ROOT_ID):
        tree[cname] = {}
        for bid, bname in parse_folder(cid):
            tree[cname][bname] = parse_folder(bid)
    json.dump(tree, open(CACHE, "w"), ensure_ascii=False)
    return tree

def norm(s): return re.sub(r"[^a-z0-9]", "", re.sub(r"\.pdf$", "", s, flags=re.I).lower())

def local_tree():
    local = {}
    for c in os.listdir(LOCAL_ROOT):
        cp = os.path.join(LOCAL_ROOT, c)
        if os.path.isdir(cp):
            local[c] = {}
            for b in os.listdir(cp):
                bp = os.path.join(cp, b)
                if os.path.isdir(bp):
                    local[c][b] = set(os.listdir(bp))
    return local

def download(fid, dest, attempts=3):
    url = f"https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t"
    for i in range(attempts):
        subprocess.run(["curl", "-sL", "--max-time", "600", "-o", dest, url])
        # validasi: PDF header / ukuran wajar, bukan halaman HTML error (quota ditolak tapi 200 OK)
        if os.path.exists(dest) and os.path.getsize(dest) > 1024:
            with open(dest, "rb") as f:
                head = f.read(1024)
            if b"<html" not in head.lower() or b"pdf" in head.lower():
                return True
        if os.path.exists(dest):
            os.remove(dest)
        time.sleep(10 * (i + 1))  # backoff: kegagalan biasanya throttle sementara
    return False

def main():
    tree = build_tree()
    local = local_tree()
    todo = []
    for cname, banks in tree.items():
        lc = next((k for k in local if norm(k) == norm(cname)), None)
        lcd = lc or cname
        for bname, files in banks.items():
            lb = next((k for k in local.get(lc, {}) if norm(k) == norm(bname)), None)
            lbd = lb or bname
            have = local.get(lc, {}).get(lb, set()) if (lc and lb) else set()
            haven = {norm(f) for f in have}
            for fid, fname in files:
                dest_dir = os.path.join(LOCAL_ROOT, lcd, lbd)
                dest = os.path.join(dest_dir, fname.replace("/", "_"))
                if os.path.exists(dest) and os.path.getsize(dest) > 1024:
                    continue
                if norm(fname) in haven:
                    continue  # sudah ada dgn nama beda
                todo.append((fid, dest))
    log(f"[{time.strftime('%H:%M:%S')}] TODO: {len(todo)} file")
    ok = fail = 0
    for i, (fid, dest) in enumerate(todo, 1):
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        if download(fid, dest):
            ok += 1
        else:
            fail += 1
            log(f"  FAIL: {dest} (id={fid})")
        if i % 25 == 0:
            log(f"  progress {i}/{len(todo)} (ok={ok} fail={fail})")
    log(f"[{time.strftime('%H:%M:%S')}] SELESAI: ok={ok} fail={fail} dari {len(todo)}")

if __name__ == "__main__":
    main()
