#!/usr/bin/env python3
"""Rename and sort recovered files by content/metadata.
EASY MODE: just double-click this file (or run it with no arguments).

Standard library only. Optional: Pillow / mutagen are NOT required.

Commands (nothing changes on disk except `backup`, `run`, `undo`):
  backup  FOLDER            full copy to FOLDER_BACKUP_<timestamp>, verified
  scan    FOLDER            counts per type, missing/wrong extensions
  dryrun  FOLDER [-n 20]    show proposed names for a random sample
  run     FOLDER            rename + sort into Photos/Videos/Music/Documents/Other
  undo    LOGFILE           reverse a run using its CSV log
"""
import argparse, collections, csv, datetime as dt, os, random, re, shutil
import struct, sys, time, zipfile
from xml.etree import ElementTree as ET

CATS = ["Photos", "Videos", "Music", "Documents", "Other"]
BACKUP_MARK = ".backup_done"
SYSTEM_DIRS = {"$RECYCLE.BIN", "System Volume Information", "FOUND.000", ".Trash-1000"}

# ---------------------------------------------------------------- detection
def sniff(path):
    """Return (kind, real_ext) from magic bytes. kind None => unknown."""
    with open(path, "rb") as f:
        h = f.read(4096)
    if h.startswith(b"\xff\xd8\xff"): return "photo", ".jpg"
    if h.startswith(b"\x89PNG\r\n\x1a\n"): return "photo", ".png"
    if h[:6] in (b"GIF87a", b"GIF89a"): return "photo", ".gif"
    if h[:2] == b"BM" and len(h) > 14: return "photo", ".bmp"
    if h[:4] in (b"II*\x00", b"MM\x00*"): return "photo", ".tif"
    if h[:4] == b"RIFF" and h[8:12] == b"WEBP": return "photo", ".webp"
    if h[4:8] == b"ftyp":
        brand = h[8:12]
        if brand in (b"heic", b"heix", b"mif1", b"heim"): return "photo", ".heic"
        if brand in (b"M4A ", b"M4B "): return "music", ".m4a"
        if brand == b"qt  ": return "video", ".mov"
        return "video", ".mp4"
    if h[:4] == b"RIFF" and h[8:12] == b"AVI ": return "video", ".avi"
    if h[:4] == b"\x1a\x45\xdf\xa3": return "video", ".mkv"
    if h[:4] == b"RIFF" and h[8:12] == b"WAVE": return "music", ".wav"
    if h[:3] == b"ID3" or (len(h) > 1 and h[0] == 0xFF and h[1] & 0xE0 == 0xE0):
        return "music", ".mp3"
    if h[:4] == b"fLaC": return "music", ".flac"
    if h[:4] == b"OggS": return "music", ".ogg"
    if h[:5] == b"%PDF-": return "doc", ".pdf"
    if h[:4] == b"PK\x03\x04":
        try:
            names = zipfile.ZipFile(path).namelist()
        except Exception:
            return None, None
        if any(n.startswith("word/") for n in names): return "doc", ".docx"
        if any(n.startswith("xl/") for n in names): return "doc", ".xlsx"
        if any(n.startswith("ppt/") for n in names): return "doc", ".pptx"
        return "other", ".zip"
    if h[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1": return "doc", ".doc"  # legacy OLE
    if h[:4] == b"Rar!": return "other", ".rar"
    if h[:6] == b"7z\xbc\xaf\x27\x1c": return "other", ".7z"
    if h[:2] == b"\x1f\x8b": return "other", ".gz"
    if h[:16] == b"SQLite format 3\x00": return "other", ".sqlite"
    if h:
        sample = h[:2048]
        if b"\x00" not in sample and _is_text(sample):
            return "doc", ".txt"
    return None, None

def _is_text(b):
    for enc in ("utf-8", "utf-16"):
        try:
            b.decode(enc, errors="strict"); return True
        except UnicodeDecodeError:
            # truncated multi-byte tail is fine
            try: b[:-3].decode(enc, errors="strict"); return True
            except UnicodeDecodeError: pass
    return False

EXT_KIND = {}
for k, exts in {
    "photo": ".jpg .jpeg .png .gif .bmp .tif .tiff .webp .heic",
    "video": ".mp4 .mov .avi .mkv .m4v .3gp .wmv",
    "music": ".mp3 .m4a .flac .wav .ogg .aac .wma",
    "doc": ".pdf .doc .docx .xls .xlsx .ppt .pptx .txt .rtf .csv .odt",
}.items():
    for e in exts.split(): EXT_KIND[e] = k
KIND_CAT = {"photo": "Photos", "video": "Videos", "music": "Music",
            "doc": "Documents", "other": "Other", None: "Other"}

def classify(path):
    kind, real_ext = sniff(path)
    cur = os.path.splitext(path)[1].lower()
    if kind is None:  # content unknown: trust extension if any
        return EXT_KIND.get(cur), cur or ""
    if kind == "doc" and real_ext == ".txt" and cur in EXT_KIND:
        return EXT_KIND[cur], cur  # csv/rtf etc. look like text
    if cur == ".jpeg" and real_ext == ".jpg": real_ext = ".jpeg"
    if cur == ".tiff" and real_ext == ".tif": real_ext = ".tiff"
    if cur == ".mov" and real_ext == ".mp4": real_ext = ".mov"
    if cur in (".m4v", ".3gp") and kind == "video": real_ext = cur
    return kind, real_ext

# ---------------------------------------------------------------- metadata
def sanitize(s, maxlen=60):
    s = re.sub(r"[\x00-\x1f]", " ", s)
    s = re.sub(r"[^\w\s\-.()]", " ", s, flags=re.UNICODE)  # keeps Arabic etc.
    s = re.sub(r"[\s_]+", " ", s).strip(" .-")
    return s[:maxlen].strip(" .-")

def fmt(d): return d.strftime("%Y-%m-%d_%H-%M-%S")

def mtime_dt(path): return dt.datetime.fromtimestamp(os.path.getmtime(path))

def exif_date(path):
    """DateTimeOriginal from JPEG/TIFF/HEIC-less EXIF (JPEG only here)."""
    with open(path, "rb") as f:
        d = f.read(131072)
    if not d.startswith(b"\xff\xd8"): return None
    i = 2
    while i + 4 < len(d):
        if d[i] != 0xFF: break
        marker, ln = d[i + 1], struct.unpack(">H", d[i + 2:i + 4])[0]
        if marker == 0xE1 and d[i + 4:i + 10] == b"Exif\x00\x00":
            t = d[i + 10:i + 2 + ln]
            return _tiff_date(t)
        i += 2 + ln
    return None

def _tiff_date(t):
    e = "<" if t[:2] == b"II" else ">"
    u16 = lambda o: struct.unpack(e + "H", t[o:o + 2])[0]
    u32 = lambda o: struct.unpack(e + "I", t[o:o + 4])[0]
    def ifd(off):
        out = {}
        for n in range(u16(off)):
            p = off + 2 + n * 12
            tag, typ, cnt, val = u16(p), u16(p + 2), u32(p + 4), u32(p + 8)
            out[tag] = (typ, cnt, val, p + 8)
        return out
    def ascii_(ent):
        typ, cnt, val, vp = ent
        o = val if cnt > 4 else vp
        return t[o:o + cnt].split(b"\x00")[0].decode("ascii", "ignore")
    try:
        i0 = ifd(u32(4))
        cands = []
        if 0x8769 in i0:
            ex = ifd(i0[0x8769][2])
            for tag in (0x9003, 0x9004):
                if tag in ex: cands.append(ascii_(ex[tag]))
        if 0x0132 in i0: cands.append(ascii_(i0[0x0132]))
        for c in cands:
            try: return dt.datetime.strptime(c.strip(), "%Y:%m:%d %H:%M:%S")
            except ValueError: continue
    except (struct.error, IndexError):
        pass
    return None

def mp4_date(path):
    """creation_time from moov/mvhd (MP4/MOV)."""
    epoch = dt.datetime(1904, 1, 1)
    with open(path, "rb") as f:
        size = os.fstat(f.fileno()).st_size
        pos = 0
        while pos + 8 <= size:
            f.seek(pos); hdr = f.read(8)
            if len(hdr) < 8: break
            ln, typ = struct.unpack(">I4s", hdr)
            hlen = 8
            if ln == 1: ln = struct.unpack(">Q", f.read(8))[0]; hlen = 16
            elif ln == 0: ln = size - pos
            if typ == b"moov":
                end, p = pos + ln, pos + hlen
                while p + 8 <= end:
                    f.seek(p); h2 = f.read(8)
                    l2, t2 = struct.unpack(">I4s", h2)
                    if t2 == b"mvhd":
                        ver = f.read(1)[0]; f.read(3)
                        ct = struct.unpack(">Q" if ver else ">I", f.read(8 if ver else 4))[0]
                        if ct == 0: return None
                        d = epoch + dt.timedelta(seconds=ct)
                        return d if d.year > 1990 else None
                    if l2 < 8: break
                    p += l2
                return None
            if ln < 8: break
            pos += ln
    return None

def _id3_text(frame):
    if not frame: return ""
    enc, b = frame[0], frame[1:]
    codec = {0: "latin-1", 1: "utf-16", 2: "utf-16-be", 3: "utf-8"}.get(enc, "latin-1")
    return b.decode(codec, "ignore").strip("\x00 ﻿").split("\x00")[0].strip()

def id3_tags(path):
    with open(path, "rb") as f:
        h = f.read(10)
        if h[:3] == b"ID3":
            ver = h[3]
            size = 0
            for b in h[6:10]: size = (size << 7) | (b & 0x7F)
            data = f.read(size)
            tags, p = {}, 0
            idlen = 4 if ver >= 3 else 3
            while p + idlen + 4 <= len(data):
                fid = data[p:p + idlen].decode("latin-1")
                if not fid.strip("\x00"): break
                if ver == 4:
                    sz = 0
                    for b in data[p + 4:p + 8]: sz = (sz << 7) | (b & 0x7F)
                    hl = 10
                elif ver == 3:
                    sz, hl = struct.unpack(">I", data[p + 4:p + 8])[0], 10
                else:
                    sz, hl = int.from_bytes(data[p + 3:p + 6], "big"), 6
                fr = data[p + hl:p + hl + sz]
                if fid in ("TPE1", "TP1"): tags["artist"] = _id3_text(fr)
                elif fid in ("TIT2", "TT2"): tags["title"] = _id3_text(fr)
                p += hl + sz
            if tags.get("title"): return tags
        f.seek(0, 2)
        if f.tell() >= 128:
            f.seek(-128, 2); t = f.read(128)
            if t[:3] == b"TAG":
                g = lambda b: b.split(b"\x00")[0].decode("latin-1", "ignore").strip()
                return {"title": g(t[3:33]), "artist": g(t[33:63])}
    return {}

def first_line(text):
    for ln in text.splitlines():
        s = sanitize(ln)
        if len(s) >= 6 and sum(c.isalpha() for c in s) >= 4:
            return s
    return ""

def office_title(path, ext):
    z = zipfile.ZipFile(path)
    try:
        root = ET.fromstring(z.read("docProps/core.xml"))
        for el in root.iter():
            if el.tag.endswith("}title") and el.text and sanitize(el.text):
                return sanitize(el.text)
    except Exception:
        pass
    if ext == ".xlsx":
        try:
            xml = z.read("xl/sharedStrings.xml").decode("utf-8", "ignore")
            texts = [re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<si>(.*?)</si>", xml, re.S)[:12]]
            s = first_line("\n".join(texts))
            if s: return s
            wb = z.read("xl/workbook.xml").decode("utf-8", "ignore")
            m = re.search(r'<sheet [^>]*name="([^"]+)"', wb)
            if m and sanitize(m.group(1)): return sanitize(m.group(1))
        except Exception:
            pass
    if ext == ".docx":
        try:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
            for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
                text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para))
                s = first_line(text)
                if s: return s
        except Exception:
            pass
    return ""

def pdf_title(path):
    with open(path, "rb") as f:
        d = f.read(1 << 20)
    m = re.search(rb"/Title\s*\((.*?)(?<!\\)\)", d, re.S)
    if m:
        raw = m.group(1)
        if raw[:2] == b"\xfe\xff": s = raw[2:].decode("utf-16-be", "ignore")
        else: s = raw.decode("latin-1", "ignore")
        s = sanitize(s)
        if len(s) >= 4 and not s.lower().startswith(("untitled", "microsoft word")):
            return s
    # uncompressed text operators, best effort
    txt = " ".join(x.decode("latin-1", "ignore") for x in re.findall(rb"\(([^()]{6,})\)\s*Tj", d))
    return first_line(txt)

def text_title(path):
    with open(path, "rb") as f:
        b = f.read(8192)
    for enc in ("utf-8-sig", "utf-16", "cp1256", "latin-1"):
        try: return first_line(b.decode(enc))
        except (UnicodeDecodeError, UnicodeError): continue
    return ""

# ---------------------------------------------------------------- naming
def propose(path):
    """Return (category, new_base, ext, source). Raises on unreadable."""
    kind, ext = classify(path)
    cat = KIND_CAT[kind]
    base, src = None, "generic"
    if kind == "photo":
        d = None
        if ext in (".jpg", ".jpeg"):
            try: d = exif_date(path)
            except Exception: d = None
        src = "exif" if d else "mtime"
        base = fmt(d or mtime_dt(path))
    elif kind == "video":
        d = None
        try: d = mp4_date(path) if ext in (".mp4", ".mov", ".m4v", ".3gp") else None
        except Exception: d = None
        src = "container" if d else "mtime"
        base = fmt(d or mtime_dt(path))
    elif kind == "music":
        tags = {}
        if ext == ".mp3":
            try: tags = id3_tags(path)
            except Exception: tags = {}
        ti, ar = sanitize(tags.get("title", ""), 50), sanitize(tags.get("artist", ""), 40)
        if ti and ar: base, src = f"{ar} - {ti}", "id3"
        elif ti: base, src = ti, "id3"
    elif kind == "doc":
        t = ""
        try:
            if ext == ".pdf": t = pdf_title(path)
            elif ext in (".docx", ".xlsx", ".pptx"): t = office_title(path, ext)
            elif ext in (".txt", ".csv", ".rtf"): t = text_title(path)
        except Exception:
            t = ""
        if t: base, src = t, "content"
    if not base:  # generic: keep stem, no guesswork
        base = sanitize(os.path.splitext(os.path.basename(path))[0], 80) or "file"
        src = "kept"
    if not ext: ext = ".bin"
    return cat, base, ext, src

def unique(dest_dir, base, ext, taken):
    n, name = 0, f"{base}{ext}"
    while os.path.lexists(os.path.join(dest_dir, name)) or \
            os.path.join(dest_dir, name).lower() in taken:
        n += 1; name = f"{base}_{n}{ext}"
    taken.add(os.path.join(dest_dir, name).lower())
    return name

# ---------------------------------------------------------------- walking
def iter_files(root):
    skip = set(CATS)
    for dp, dn, fn in os.walk(root):
        rel = os.path.relpath(dp, root)
        dn[:] = [d for d in dn if d not in SYSTEM_DIRS]
        if rel == ".": dn[:] = [d for d in dn if d not in skip]  # already sorted
        for n in sorted(fn):
            if n == BACKUP_MARK or n.startswith("rename_log_"): continue
            yield os.path.join(dp, n)

# ---------------------------------------------------------------- commands
def cmd_backup(a):
    src = os.path.abspath(a.folder)
    files = list(iter_files(src))
    total = sum(os.path.getsize(p) for p in files)
    dest = os.path.join(os.path.abspath(a.dest), f"BACKUP_{time.strftime('%Y%m%d_%H%M%S')}")
    if os.path.commonpath([src, dest]) == src:
        sys.exit("Backup destination must be OUTSIDE the source folder (use another drive).")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    free = shutil.disk_usage(os.path.dirname(dest)).free
    print(f"{len(files)} files, {total/1e9:.2f} GB; free space {free/1e9:.2f} GB")
    if free < total * 1.05:
        sys.exit("Not enough free disk space for a full backup. Aborting.")
    shutil.copytree(src, dest)
    copied = list(iter_files(dest))
    ctotal = sum(os.path.getsize(p) for p in copied)
    if len(copied) != len(files) or ctotal != total:
        sys.exit(f"BACKUP VERIFY FAILED: {len(copied)} files/{ctotal} bytes vs {len(files)}/{total}")
    with open(os.path.join(src, BACKUP_MARK), "w") as f: f.write(dest + "\n")
    print(f"Backup verified: {dest}")

def cmd_scan(a):
    root = os.path.abspath(a.folder)
    cnt, noext, wrong, unreadable = collections.Counter(), [], [], []
    for p in iter_files(root):
        try:
            kind, ext = classify(p)
        except OSError as e:
            unreadable.append((p, str(e))); continue
        cnt[KIND_CAT[kind]] += 1
        cur = os.path.splitext(p)[1].lower()
        if not cur: noext.append(p)
        elif ext and cur != ext and EXT_KIND.get(cur) != kind: wrong.append(p)
    print(f"Total files: {sum(cnt.values())}")
    for c in CATS: print(f"  {c:10s} {cnt[c]}")
    print(f"Missing extension: {len(noext)}   Wrong extension: {len(wrong)}   Unreadable: {len(unreadable)}")
    for p in noext[:5]: print("   e.g. no ext:", os.path.basename(p))

def plan(root, files):
    taken, rows = set(), []
    for p in files:
        try:
            cat, base, ext, src = propose(p)
            dest_dir = os.path.join(root, cat)
            name = unique(dest_dir, base, ext, taken)
            rows.append((p, os.path.join(dest_dir, name), src, None))
        except Exception as e:
            rows.append((p, None, "skipped", f"{type(e).__name__}: {e}"))
    return rows

def cmd_dryrun(a):
    root = os.path.abspath(a.folder)
    files = list(iter_files(root))
    random.Random(a.seed).shuffle(files)
    rows = plan(root, files[:a.n])
    print(f"DRY RUN - {len(rows)} sample files, nothing changed\n")
    for old, new, src, err in rows:
        o = os.path.relpath(old, root)
        print(f"{o}\n   -> {os.path.relpath(new, root) if new else 'SKIP ('+err+')'}   [{src}]")

def cmd_run(a):
    root = os.path.abspath(a.folder)
    if not os.path.exists(os.path.join(root, BACKUP_MARK)):
        sys.exit("No verified backup found. Run `backup` first.")
    files = list(iter_files(root))
    rows = plan(root, files)
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = skipped = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, src, err in rows:
            if new is None:
                w.writerow([old, "", "skipped", "", err]); skipped += 1; continue
            try:
                os.makedirs(os.path.dirname(new), exist_ok=True)
                if os.path.lexists(new): raise FileExistsError(new)  # never overwrite
                os.rename(old, new)
                w.writerow([old, new, "renamed", src, ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", f"{type(e).__name__}: {e}"]); skipped += 1
            lf.flush()
    print(f"Renamed: {ok}\nSkipped: {skipped} (listed in log)\nLog: {logp}")

def cmd_office(a):
    """Rename ONLY Word/Excel files, in place. No copy, no moving, no extra space."""
    root = os.path.abspath(a.folder)
    cands, legacy, seen = [], 0, 0
    print("Scanning files (progress shown every 1000)...", flush=True)
    for p in iter_files(root):
        seen += 1
        if seen % 1000 == 0: print(f"  scanned {seen} files, {len(cands)} Word/Excel so far", flush=True)
        try:
            kind, ext = classify(p)
        except OSError:
            continue
        if ext in (".docx", ".xlsx"): cands.append(p)
        elif kind == "doc" and ext == ".doc": legacy += 1
    print(f"Word/Excel files found: {len(cands)} (old-format .doc/.xls left untouched: {legacy})")
    taken, rows = set(), []
    for i, p in enumerate(cands, 1):
        if i % 500 == 0: print(f"  reading titles: {i}/{len(cands)}", flush=True)
        try:
            _, base, ext, src = propose(p)
            if src == "kept": rows.append((p, None, "no title found", None)); continue
            name = unique(os.path.dirname(p), base, ext, taken)
            new = os.path.join(os.path.dirname(p), name)
            if os.path.normcase(new) == os.path.normcase(p): continue
            rows.append((p, new, src, None))
        except Exception as e:
            rows.append((p, None, "skipped", f"{type(e).__name__}: {e}"))
    todo = [r for r in rows if r[1]]
    if not a.apply:
        random.Random(1).shuffle(todo)
        print(f"\nPREVIEW ONLY (nothing changed). {len(todo)} would be renamed; showing {min(a.n, len(todo))}:\n")
        for old, new, src, _ in todo[:a.n]:
            print(f"{os.path.relpath(old, root)}\n   -> {os.path.basename(new)}   [{src}]")
        print(f"\n{len(rows)-len(todo)} would be left as they are (no readable title).")
        if not (todo and sys.stdin.isatty()): return
        if input(f"\nType YES to rename all {len(todo)} files now (anything else = stop): ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = skipped = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, src, err in rows:
            if not new: w.writerow([old, "", "skipped", "", src if not err else err]); skipped += 1; continue
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", src, ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", f"{type(e).__name__}: {e}"]); skipped += 1
            lf.flush()
    print(f"Renamed: {ok}\nLeft as-is/skipped: {skipped}\nUndo log: {logp}")

def cmd_media(a):
    """Rename photos/videos IN PLACE by date taken. Only when a real date is inside the file."""
    root = os.path.abspath(a.folder)
    print("Scanning files (progress shown every 1000)...", flush=True)
    rows, seen, nodate, stats = [], 0, 0, collections.Counter()
    taken = set()
    for p in iter_files(root):
        seen += 1
        if seen % 1000 == 0: print(f"  scanned {seen} files, {len(rows)} to rename so far", flush=True)
        try:
            kind, _ = classify(p)
            if kind not in ("photo", "video"): continue
            _, base, ext, src = propose(p)
            stats[kind] += 1
            if src not in ("exif", "container"):   # mtime of recovered files is unreliable
                nodate += 1; rows.append((p, None, "no date inside file", None)); continue
            stem = os.path.splitext(os.path.basename(p))[0]
            if re.fullmatch(re.escape(base) + r"(_\d+)?", stem):  # already named this way
                taken.add(p.lower()); continue
            name = unique(os.path.dirname(p), base, ext, taken)
            new = os.path.join(os.path.dirname(p), name)
            rows.append((p, new, src, None))
        except Exception as e:
            rows.append((p, None, "skipped", f"{type(e).__name__}: {e}"))
    todo = [r for r in rows if r[1]]
    print(f"\nPhotos: {stats['photo']}  Videos: {stats['video']}")
    print(f"With a real date inside: {len(todo)}   No date inside (left unchanged): {len(rows)-len(todo)}")
    if not a.apply:
        random.Random(1).shuffle(todo)
        print(f"\nPREVIEW ONLY. Showing {min(a.n, len(todo))} examples:\n")
        for old, new, src, _ in todo[:a.n]:
            print(f"{os.path.relpath(old, root)}\n   -> {os.path.basename(new)}   [{src}]")
        if not (todo and sys.stdin.isatty()): return
        if input(f"\nType YES (capitals) to rename all {len(todo)} files now: ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = skipped = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, src, err in rows:
            if not new: w.writerow([old, "", "skipped", "", src if not err else err]); skipped += 1; continue
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", src, ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", f"{type(e).__name__}: {e}"]); skipped += 1
            lf.flush()
    print(f"Renamed: {ok}\nLeft as-is/skipped: {skipped}\nUndo log: {logp}")

_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_NORM = str.maketrans({"ة": "ه", "ى": "ي", "أ": "ا", "إ": "ا", "آ": "ا", "ـ": None})
_CASE_KEY = r"(?:القضيه|قضيه|الدعوي|الدعوه)"
_CASE_RE = re.compile(_CASE_KEY + r"[^\d]{0,25}?(\d+(?:\s*[/\\\-]\s*\d+){0,2})"
                      r"(?:\s*(?:لسنه|لعام|عام|سنه|لسنة)\s*(\d{2,4}))?")

def docx_text(path, limit=200000):
    z = zipfile.ZipFile(path)
    parts = ["word/document.xml"] + [n for n in z.namelist() if re.match(r"word/header\d*\.xml", n)]
    out = []
    for n in parts:
        try: xml = z.read(n).decode("utf-8", "ignore")
        except KeyError: continue
        for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
            out.append("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)))
        if sum(map(len, out)) > limit: break
    return " ".join(out)

def case_number(path):
    t = re.sub(r"[\u064B-\u065F]", "", docx_text(path)).translate(_AR_DIGITS).translate(_NORM)
    t = re.sub(r"\s+", " ", t)
    for m in _CASE_RE.finditer(t):
        num = re.sub(r"\s*[/\\]\s*|\s*-\s*", "-", m.group(1).strip())
        parts = num.split("-")
        if len(parts) == 3 and 1 <= int(parts[0]) <= 31 and 1 <= int(parts[1]) <= 12 and len(parts[2]) == 4:
            continue                      # looks like a date (27/3/2014), not a case number
        if len(parts) == 1 and not m.group(2) and len(parts[0]) < 2:
            continue                      # lone single digit
        if m.group(2): num += "-" + m.group(2)
        return num, t[max(0, m.start() - 10):m.end() + 5]
    return None, None

def cmd_cases(a):
    """Rename Word files to their case number (رقم القضية), in place."""
    root = os.path.abspath(a.folder)
    print("Scanning Word files (progress every 500)...", flush=True)
    rows, taken, seen, nomatch, already = [], set(), 0, 0, 0
    errs, samples, qcontext = collections.Counter(), [], []
    for p in iter_files(root):
        if not p.lower().endswith(".docx"): continue
        seen += 1
        if seen % 500 == 0: print(f"  checked {seen}, matched {len([r for r in rows if r[1]])}", flush=True)
        try:
            num, snip = case_number(p)
        except Exception as e:
            errs[f"{type(e).__name__}: {str(e)[:60]}"] += 1
            rows.append((p, None, f"{type(e).__name__}: {e}")); continue
        if not num:
            nomatch += 1
            if len(samples) < 3 or len(qcontext) < 4:
                t = re.sub(r"\s+", " ", docx_text(p))
                if len(samples) < 3 and t.strip(): samples.append((p, t[:200]))
                i = t.find("قض")
                if i >= 0 and len(qcontext) < 4: qcontext.append((p, t[max(0, i - 40):i + 80]))
            continue
        base = a.prefix + num
        stem = os.path.splitext(os.path.basename(p))[0]
        if re.fullmatch(re.escape(base) + r"(_\d+)?", stem): taken.add(p.lower()); already += 1; continue
        name = unique(os.path.dirname(p), base, ".docx", taken)
        rows.append((p, os.path.join(os.path.dirname(p), name), snip))
    todo = [r for r in rows if r[1]]
    print(f"\nWord files checked: {seen}\nCase number found: {len(todo)}   Not found (left unchanged): {nomatch}   Already named: {already}   Errors: {sum(errs.values())}")
    for k, v in errs.most_common(4): print(f"  error x{v}: {k}")
    if not todo:
        print("\n--- DIAGNOSTIC: text from files with no match ---")
        for p, t in samples: print(f"{os.path.basename(p)}: {t}")
        print("--- places where the word 'قض' appears ---")
        for p, t in qcontext: print(f"{os.path.basename(p)}: ...{t}...")
    if not a.apply:
        random.Random(1).shuffle(todo)
        print(f"\nPREVIEW ONLY. {min(a.n, len(todo))} examples (old -> new  [text matched]):\n")
        for old, new, snip in todo[:a.n]:
            print(f"{os.path.relpath(old, root)}\n   -> {os.path.basename(new)}   [{snip}]")
        if not (todo and sys.stdin.isatty()): return
        if input(f"\nType YES (capitals) to rename all {len(todo)} files now: ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = skipped = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, note in rows:
            if not new: w.writerow([old, "", "skipped", "", note]); skipped += 1; continue
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", "case-number", ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", f"{type(e).__name__}: {e}"]); skipped += 1
            lf.flush()
    print(f"Renamed: {ok}\nSkipped: {skipped}\nUndo log: {logp}")

def cmd_previews(a):
    """Rename preview PNGs (made by word_previews.ps1) to the case number found in their Word file."""
    pdir, root = os.path.abspath(a.previews), os.path.abspath(a.docx_root)
    if a.unknown == "auto": a.unknown = "قضية رقم ؟"
    key = lambda rel: re.sub(r'[\\/:*?"<>|]', " - ", rel) + ".png"
    cur = {}                                    # preview-name -> current docx path
    print("Step 1/3: listing Word files on the drive...", flush=True)
    for p in iter_files(root):
        if p.lower().endswith(".docx"):
            cur[key(os.path.relpath(p, root))] = p
            if len(cur) % 500 == 0: print(f"  listed {len(cur)} Word files", flush=True)
    print(f"  listed {len(cur)} Word files in total", flush=True)
    print("Step 2/3: reading earlier rename logs...", flush=True)
    moved = {}                                  # follow earlier renames recorded in logs
    existing = set(cur.values())
    logs = [] if a.no_logs else sorted(f for f in os.listdir(root) if f.startswith("rename_log_") and f.endswith(".csv"))
    if a.no_logs: print("  (skipped on request)", flush=True)
    for lg in logs:
        print(f"  reading {lg} ...", flush=True)
        with open(os.path.join(root, lg), newline="", encoding="utf-8-sig") as f:
            n = 0
            for r in csv.DictReader(f):
                if r["status"] == "renamed": moved[r["old_path"]] = r["new_path"]; n += 1
        print(f"  log {lg}: {n} renames", flush=True)
    resolved = 0
    for rl in list(moved):                      # old name -> final name (guarded against loops)
        d, hops = moved[rl], 0
        while d in moved and d != moved[d] and hops < 20: d = moved[d]; hops += 1
        if d in existing:
            cur.setdefault(key(os.path.relpath(rl, root)), d); resolved += 1
    print(f"  matched {resolved} old names to current files", flush=True)
    rows, taken, keep, nodoc, seen, nocase = [], set(), 0, 0, 0, []
    print("Step 3/3: reading each document's text (progress every 100)...", flush=True)
    for png in sorted(os.listdir(pdir)):
        if not png.lower().endswith(".png"): continue
        seen += 1
        if seen % 100 == 0: print(f"  checked {seen} pictures, to rename so far: {len(rows)}", flush=True)
        src = cur.get(png)
        if not src: nodoc += 1; continue
        try: num, snip = case_number(src)
        except Exception: num = None
        if not num:
            keep += 1
            if len(nocase) < 6:
                try: nocase.append((png, re.sub(r"\s+", " ", docx_text(src))[:160]))
                except Exception: pass
            if a.unknown:
                orig = re.sub(r"\.docx\.png$", "", png, flags=re.I)
                ubase = a.unknown if a.plain else f"{a.unknown} - {orig}"
                ubase = ubase[:150]
                if png[:-4] == ubase or re.fullmatch(re.escape(ubase) + r"_\d+", png[:-4]): continue
                rows.append((os.path.join(pdir, png), os.path.join(pdir, unique(pdir, ubase, ".png", taken)), "no case number"))
            continue
        base = a.prefix + num
        if re.fullmatch(re.escape(base) + r"_\d+|" + re.escape(base), png[:-4]): continue
        rows.append((os.path.join(pdir, png), os.path.join(pdir, unique(pdir, base, ".png", taken)), snip))
    print(f"\nPictures: {seen}\nWill be renamed by case number: {len(rows)}\nKept as they are (no case number): {keep}\nCould not match to a Word file: {nodoc}")
    if not a.apply:
        random.Random(1).shuffle(rows)
        print(f"\nPREVIEW ONLY. {min(a.n, len(rows))} examples:\n")
        for old, new, snip in rows[:a.n]: print(f"{os.path.basename(old)}\n   -> {os.path.basename(new)}   [{snip}]")
        if nocase:
            print("\n--- pictures with NO case number found (start of their text) ---")
            for n_, t_ in nocase: print(f"{n_}\n   {t_}")
        if not (rows and sys.stdin.isatty()): return
        if input(f"\nType YES (capitals) to rename all {len(rows)} pictures now: ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(pdir, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, _ in rows:
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", "case-number", ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", str(e)])
    print(f"Renamed: {ok}\nUndo log: {logp}")

def cmd_undo(a):
    n = 0
    with open(a.log, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["status"] != "renamed": continue
            if os.path.exists(r["new_path"]) and not os.path.lexists(r["old_path"]):
                os.makedirs(os.path.dirname(r["old_path"]), exist_ok=True)
                os.rename(r["new_path"], r["old_path"]); n += 1
    print(f"Restored {n} files to original names/locations")

def _ask(q, default=""):
    v = input(f"{q} [{default}]: ").strip().strip('"')
    return v or default

def _pick_backup_drive(src, need):
    import string
    best = None
    for L in string.ascii_uppercase:
        d = f"{L}:\\"
        if not os.path.exists(d) or os.path.splitdrive(src)[0].upper() == f"{L}:": continue
        try: free = shutil.disk_usage(d).free
        except OSError: continue
        if free > need * 1.05 and (best is None or free > best[1]): best = (d, free)
    return best[0] if best else None

def wizard():
    """Double-click mode: guided, asks before each risky step."""
    try:
        print("=== Recovered files: rename & sort ===\n")
        folder = os.path.abspath(_ask("Folder with recovered files", r"I:\Recovery"))
        if not os.path.isdir(folder): sys.exit(f"Folder not found: {folder}")
        ns = argparse.Namespace(folder=folder)
        print("\n[1/4] Scanning..."); cmd_scan(ns)
        files = list(iter_files(folder))
        need = sum(os.path.getsize(f) for f in files)
        drive = _pick_backup_drive(folder, need)
        dest = _ask(f"\n[2/4] Backup needs {need/1e9:.1f} GB. Backup folder (must be on ANOTHER drive)",
                    os.path.join(drive, "Recovery_Backup") if drive else "")
        if not dest: sys.exit("No other drive with enough space found. Connect one and try again.")
        if _ask("Start backup? (yes/no)", "yes").lower() != "yes": sys.exit("Stopped. Nothing changed.")
        cmd_backup(argparse.Namespace(folder=folder, dest=dest))
        print("\n[3/4] Preview of 20 proposed names (nothing renamed yet):\n")
        cmd_dryrun(argparse.Namespace(folder=folder, n=20, seed=1))
        if _ask("\nNames look OK? Rename ALL files now? (yes/no)", "no").lower() != "yes":
            sys.exit("Stopped. Nothing was renamed.")
        print("\n[4/4] Renaming..."); cmd_run(ns)
        print("\nDone. To undo, run: python recover_rename.py undo <log file above>")
    except SystemExit as e:
        if e.code: print(e.code)
    except Exception as e:
        print("ERROR:", e)
    input("\nPress Enter to close...")

def main():
    if len(sys.argv) == 1: return wizard()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in (("backup", cmd_backup), ("scan", cmd_scan), ("dryrun", cmd_dryrun), ("run", cmd_run)):
        p = sp.add_parser(name); p.add_argument("folder"); p.set_defaults(fn=fn)
        if name == "backup": p.add_argument("--dest", required=True, help="folder on ANOTHER drive")
        if name == "dryrun":
            p.add_argument("-n", type=int, default=20); p.add_argument("--seed", type=int, default=1)
    p = sp.add_parser("office"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("-n", type=int, default=20); p.set_defaults(fn=cmd_office)
    p = sp.add_parser("media"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("-n", type=int, default=20); p.set_defaults(fn=cmd_media)
    p = sp.add_parser("cases"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("-n", type=int, default=20); p.add_argument("--prefix", default="قضية "); p.set_defaults(fn=cmd_cases)
    p = sp.add_parser("previews"); p.add_argument("previews"); p.add_argument("docx_root")
    p.add_argument("--apply", action="store_true"); p.add_argument("-n", type=int, default=20)
    p.add_argument("--prefix", default="قضية ")
    p.add_argument("--unknown", default="", help="name for pictures with no case number, e.g. 'قضية رقم ؟'")
    p.add_argument("--plain", action="store_true", help="with --unknown: do not append the old title")
    p.add_argument("--no-logs", action="store_true", help="skip reading earlier rename logs")
    p.set_defaults(fn=cmd_previews)
    p = sp.add_parser("undo"); p.add_argument("log"); p.set_defaults(fn=cmd_undo)
    a = ap.parse_args(); a.fn(a)

if __name__ == "__main__":
    main()
