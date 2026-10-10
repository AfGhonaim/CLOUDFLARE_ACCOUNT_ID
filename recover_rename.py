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
import argparse, collections, csv, datetime as dt, html, os, random, re, shutil, zlib
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

def _case_from_text(t):
    t = re.sub(r"[\u064B-\u065F]", "", t).translate(_AR_DIGITS).translate(_NORM)
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

def case_number(path):
    return _case_from_text(docx_text(path))

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

def _picture_map(pdir, root):
    """Return [(docx_path, png_path)] linking each preview picture to its CURRENT Word file."""
    key = lambda rel: re.sub(r'[\\/:*?"<>|]', " - ", rel) + ".png"
    cur = {}
    for p in iter_files(root):
        if p.lower().endswith(".docx"): cur[key(os.path.relpath(p, root))] = p
    existing = set(cur.values())
    moved = {}
    for lg in sorted(f for f in os.listdir(root) if f.startswith("rename_log_") and f.endswith(".csv")):
        with open(os.path.join(root, lg), newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if r["status"] == "renamed": moved[r["old_path"]] = r["new_path"]
    for rl in list(moved):
        d, hops = moved[rl], 0
        while d in moved and d != moved[d] and hops < 20: d = moved[d]; hops += 1
        if d in existing: cur.setdefault(key(os.path.relpath(rl, root)), d)
    out, used = [], set()
    for png in sorted(os.listdir(pdir)):
        if not png.lower().endswith(".png"): continue
        cand = [png]
        m = re.match(r"^قضية رقم ؟ - (.*)\.png$", png)       # picture renamed because no case number
        if m: cand += [m.group(1) + ".docx.png", m.group(1) + ".png"]
        for c in cand:
            d = cur.get(c)
            if d and d not in used:
                used.add(d); out.append((d, os.path.join(pdir, png))); break
    return out

_JPG_PS = r"""
Add-Type -AssemblyName System.Drawing
$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]80)
Import-Csv -LiteralPath $args[0] -Encoding UTF8 | ForEach-Object {
  try {
    $img = [System.Drawing.Image]::FromFile($_.png)
    $sc = 256.0 / [Math]::Max($img.Width, $img.Height)
    $w = [Math]::Max(1, [int]($img.Width * $sc)); $h = [Math]::Max(1, [int]($img.Height * $sc))
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.DrawImage($img, 0, 0, $w, $h)
    $bmp.Save($_.jpg, $codec, $ep)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose()
  } catch { }
}
"""

_WORD_PS = r"""
Add-Type -AssemblyName System.Drawing
$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]80)
$rows = @(Import-Csv -LiteralPath $args[0] -Encoding UTF8)
try { $word = New-Object -ComObject Word.Application } catch { Write-Host "Microsoft Word could not be started."; exit 1 }
$word.Visible = $false; $word.DisplayAlerts = 0; $word.AutomationSecurity = 3
$n = 0
foreach ($r in $rows) {
  $n++
  if ((Test-Path -LiteralPath $r.jpg) -and ((Get-Item -LiteralPath $r.jpg).Length -gt 200)) { continue }
  $doc = $null
  try {
    $doc = $word.Documents.Open($r.docx, $false, $true, $false)
    $doc.ActiveWindow.View.Type = 3
    $bytes = $doc.ActiveWindow.ActivePane.Pages.Item(1).EnhMetaFileBits
    $ms = New-Object System.IO.MemoryStream (, $bytes)
    $mf = New-Object System.Drawing.Imaging.Metafile($ms)
    $sc = 256.0 / [Math]::Max($mf.Width, $mf.Height)
    $w = [Math]::Max(1, [int]($mf.Width * $sc)); $h = [Math]::Max(1, [int]($mf.Height * $sc))
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White); $g.InterpolationMode = 'HighQualityBicubic'
    $g.DrawImage($mf, 0, 0, $w, $h)
    $bmp.Save($r.jpg, $codec, $ep)
    $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); $ms.Dispose()
  } catch { } finally { if ($doc) { try { $doc.Close($false) } catch { } } }
  if ($n % 25 -eq 0) { Write-Host ("  pictures made: {0}/{1}" -f $n, $rows.Count) }
}
$word.Quit()
"""

_THUMB_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"

def embed_thumbnail(docx, jpg_bytes):
    """Add docProps/thumbnail.jpeg to a .docx without touching its content. Atomic; returns status."""
    tmp = docx + ".thumbtmp"
    with zipfile.ZipFile(docx) as zin:
        names = zin.namelist()
        if "_rels/.rels" not in names or "[Content_Types].xml" not in names: return "skipped: not a normal docx"
        if "docProps/thumbnail.jpeg" in names and zin.read("docProps/thumbnail.jpeg") == jpg_bytes: return "already"
        rels = zin.read("_rels/.rels").decode("utf-8")
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        if _THUMB_REL in rels:
            rels = re.sub(r'(<Relationship\b[^>]*Type="' + re.escape(_THUMB_REL) + r'"[^>]*Target=")[^"]*(")',
                          r"\1docProps/thumbnail.jpeg\2", rels)
            rels = re.sub(r'(<Relationship\b[^>]*Target=")[^"]*("[^>]*Type="' + re.escape(_THUMB_REL) + r'")',
                          r"\1docProps/thumbnail.jpeg\2", rels)
        else:
            rels = rels.replace("</Relationships>", f'<Relationship Id="rIdThumb1" Type="{_THUMB_REL}" Target="docProps/thumbnail.jpeg"/></Relationships>')
        if not re.search(r'Extension="jpe?g"', ct, re.I):
            ct = re.sub(r"(<Types\b[^>]*>)", r'\1<Default Extension="jpeg" ContentType="image/jpeg"/>', ct, count=1)
        try:
            with zipfile.ZipFile(tmp, "w") as zout:
                for item in zin.infolist():
                    if item.filename in ("_rels/.rels", "[Content_Types].xml", "docProps/thumbnail.jpeg"): continue
                    zout.writestr(item, zin.read(item.filename))
                zout.writestr("_rels/.rels", rels.encode("utf-8"), zipfile.ZIP_DEFLATED)
                zout.writestr("[Content_Types].xml", ct.encode("utf-8"), zipfile.ZIP_DEFLATED)
                zout.writestr("docProps/thumbnail.jpeg", jpg_bytes, zipfile.ZIP_STORED)
            with zipfile.ZipFile(tmp) as chk:                       # verify before replacing
                if chk.testzip() is not None or len(chk.namelist()) != len(set(names) | {"docProps/thumbnail.jpeg"}):
                    raise ValueError("verification failed")
        except Exception as e:
            if os.path.exists(tmp): os.remove(tmp)
            return f"failed: {e}"
    st = os.stat(docx)
    os.replace(tmp, docx)
    os.utime(docx, (st.st_atime, st.st_mtime))                      # keep original dates
    return "added"

def cmd_thumbs(a):
    """Give every Word file a page-1 thumbnail (stored inside the .docx; the file stays a normal Word file)."""
    pdir, root = os.path.abspath(a.previews), os.path.abspath(a.docx_root)
    jdir = os.path.join(os.environ.get("TEMP", pdir), "thumbs_jpg"); os.makedirs(jdir, exist_ok=True)
    if a.from_pictures:
        print("Linking pictures to Word files...", flush=True)
        pairs = _picture_map(pdir, root); print(f"  linked: {len(pairs)}", flush=True)
        todo = [(d, pg) for d, pg in pairs]
    else:
        print("Listing Word files (those without a thumbnail)...", flush=True)
        todo = []
        for q in iter_files(root):
            if not q.lower().endswith(".docx"): continue
            try:
                with zipfile.ZipFile(q) as z:
                    has = any(n.lower().startswith("docprops/thumbnail") for n in z.namelist())
            except Exception:
                continue
            if a.force or not has: todo.append((q, None))
        print(f"  Word files needing a thumbnail: {len(todo)}", flush=True)
    if not a.all: todo = todo[:a.limit]; print(f"TEST MODE: only the first {len(todo)} files. Add --all for everything.")
    mp = os.path.join(jdir, "map.csv")
    with open(mp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["docx", "png", "jpg"])
        for i, (d, png) in enumerate(todo): w.writerow([d, png or "", os.path.join(jdir, f"{i}.jpg")])
    if not a.skip_convert:
        import subprocess
        if a.from_pictures:
            print("Making small JPEG copies of the pictures (Windows)...", flush=True)
            ps = os.path.join(jdir, "conv.ps1"); open(ps, "w", encoding="utf-8-sig").write(_JPG_PS)
        else:
            print("Opening each file in Word to make its page-1 picture (this takes a while)...", flush=True)
            ps = os.path.join(jdir, "word.ps1"); open(ps, "w", encoding="utf-8-sig").write(_WORD_PS)
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps, mp], check=False)
    logp = os.path.join(pdir, f"thumbs_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    cnt = collections.Counter()
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["docx", "picture", "status"])
        for i, (d, png) in enumerate(todo, 1):
            jp = os.path.join(jdir, f"{i-1}.jpg")
            try:
                if not os.path.exists(jp) or os.path.getsize(jp) < 200: st = "skipped: no picture made"
                else: st = embed_thumbnail(d, open(jp, "rb").read())
            except Exception as e:
                st = f"failed: {type(e).__name__}: {e}"
            cnt[st.split(":")[0]] += 1; w.writerow([d, png or "", st]); lf.flush()
            if i % 100 == 0: print(f"  embedded {i}/{len(todo)}  added: {cnt['added']}", flush=True)
    print(f"\nAdded thumbnails: {cnt['added']}   Already had: {cnt['already']}   Skipped/failed: {len(todo)-cnt['added']-cnt['already']}\nLog: {logp}")

def _xlsx_text(path, limit=300):
    z = zipfile.ZipFile(path)
    try:
        xml = z.read("xl/sharedStrings.xml").decode("utf-8", "ignore")
        return html.unescape("\n".join(re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<si>(.*?)</si>", xml, re.S)[:limit]))
    except KeyError:                                   # some writers store text inline in the sheet
        try: xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
        except KeyError: return ""
        return html.unescape("\n".join(re.findall(r"<t[^>]*>([^<]+)</t>", xml)[:limit]))

def _pdf_text_std(path, content=False):
    """Best-effort PDF text with the standard library: Latin text and the /Title only."""
    with open(path, "rb") as f: data = f.read(6_000_000)
    out = []
    m = re.search(rb"/Title\s*\((.*?)(?<!\\)\)", data, re.S)
    if m:
        raw = m.group(1)
        t = raw[2:].decode("utf-16-be", "ignore") if raw[:2] == b"\xfe\xff" else raw.decode("latin-1", "ignore")
        t = re.sub(r"^(microsoft word|microsoft excel|adobe acrobat)\s*-\s*", "", t.strip(), flags=re.I)
        t = re.sub(r"\.(docx?|xlsx?|pptx?|txt|indd)$", "", t, flags=re.I)
        if len(t) >= 4 and not re.match(r"(untitled|document\d*|scan|new document|title)", t, re.I): out.append(t)
    if not content: return "\n".join(out)          # built-in reader drops letters: use only the PDF's own title
    import base64
    for sm in re.finditer(rb"stream\r?\n(.*?)\r?\n?endstream", data, re.S):
        raw = sm.group(1)
        try: dec = zlib.decompress(raw)
        except Exception:
            try:
                r85 = raw.strip()
                r85 = r85[:-2] if r85.endswith(b"~>") else r85
                dec = zlib.decompress(base64.a85decode(r85))
            except Exception: dec = raw
        if b"BT" not in dec: continue
        for t1, t2 in re.findall(rb"\(((?:[^()\\]|\\.)*)\)\s*Tj|\[((?:[^\]])*)\]\s*TJ", dec, re.S):
            parts = [t1] if t1 else re.findall(rb"\(((?:[^()\\]|\\.)*)\)", t2)
            line = b"".join(parts).decode("latin-1", "ignore")
            if line and sum(c.isascii() for c in line) >= 0.9 * len(line): out.append(line)
        if sum(map(len, out)) > 4000: break
    return "\n".join(out)

def _pdf_text(path):
    try:
        import pypdf
        r = pypdf.PdfReader(path)
        if r.is_encrypted: r.decrypt("")
        txt = "\n".join((pg.extract_text() or "") for pg in r.pages[:3])
        if txt.strip(): return txt, "pypdf"
    except KeyboardInterrupt:
        raise
    except BaseException:                     # library missing or broken: fall back to the built-in reader
        pass
    return _pdf_text_std(path, content=False), "std"

def cmd_content(a):
    """Rename PDF and Excel (.xlsx) files in place by content: case number first, else title/first line."""
    root = os.path.abspath(a.folder)
    print("Scanning (progress every 500 files)...", flush=True)
    rows, taken, seen, found, keep, errs = [], set(), 0, 0, 0, collections.Counter()
    for p in iter_files(root):
        seen += 1
        if seen % 500 == 0: print(f"  scanned {seen} files, PDF/Excel found: {found}, to rename: {len([r for r in rows if r[1]])}", flush=True)
        try:
            kind, ext = classify(p)
            if ext not in (".pdf", ".xlsx"): continue
            found += 1
            if ext == ".pdf": text, how = _pdf_text(p)
            else: text, how = _xlsx_text(p), "xlsx"
            if ext == ".pdf": text = _fix_dir(text)
            num, snip = _case_from_text(text)
            if num: base, src = a.prefix + num, "case number"
            else:
                t = first_line(text)
                if ext == ".xlsx" and not t: t = office_title(p, ".xlsx")
                if not t: keep += 1; continue
                base, src = t, f"title ({how})"
            stem = os.path.splitext(os.path.basename(p))[0]
            if re.fullmatch(re.escape(base) + r"(_\d+)?", stem): taken.add(p.lower()); continue
            name = unique(os.path.dirname(p), base, ext, taken)
            rows.append((p, os.path.join(os.path.dirname(p), name), src))
        except Exception as e:
            errs[type(e).__name__] += 1
    todo = [r for r in rows if r[1]]
    by = collections.Counter(r[2] for r in todo)
    print(f"\nPDF/Excel files found: {found}\nWill be renamed: {len(todo)}  {dict(by)}\nNo readable content (left unchanged): {keep}   Errors: {sum(errs.values())}")
    if not a.apply:
        random.Random(1).shuffle(todo)
        print(f"\nPREVIEW ONLY. {min(a.n, len(todo))} examples:\n")
        for old, new, src in todo[:a.n]: print(f"{os.path.relpath(old, root)}\n   -> {os.path.basename(new)}   [{src}]")
        if not (todo and sys.stdin.isatty()): return
        if input(f"\nType YES (capitals) to rename all {len(todo)} files now: ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, src in todo:
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", src, ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", str(e)])
            lf.flush()
    print(f"Renamed: {ok}\nUndo log: {logp}")

def _xlsx_peek(path, nrows=4, ncells=6, maxlen=40):
    z = zipfile.ZipFile(path)
    sheets = [html.unescape(m) for m in re.findall(r'<sheet [^>]*name="([^"]*)"', z.read("xl/workbook.xml").decode("utf-8", "ignore"))]
    try:
        sst = [html.unescape(re.sub(r"<[^>]+>", "", m)) for m in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode("utf-8", "ignore"), re.S)]
    except KeyError:
        sst = []
    names = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
    xml = z.read(sorted(names)[0]).decode("utf-8", "ignore") if names else ""
    rows = []
    for rm in re.finditer(r"<row\b[^>]*>(.*?)</row>", xml, re.S):
        vals = []
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", rm.group(1), re.S):
            attrs, body = cm.group(1), cm.group(2) or ""
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            if 't="s"' in attrs and v and v.group(1).isdigit() and int(v.group(1)) < len(sst): vals.append(sst[int(v.group(1))])
            elif "inlineStr" in attrs: vals.append(html.unescape(re.sub(r"<[^>]+>", "", body)))
            elif v: vals.append(v.group(1))
        vals = [re.sub(r"\s+", " ", x).strip()[:maxlen] for x in vals if x.strip()]
        if vals: rows.append(vals[:ncells])
        if len(rows) >= nrows: break
    return sheets, rows

def cmd_xlsxpeek(a):
    root = os.path.abspath(a.folder)
    files = [p for p in iter_files(root) if p.lower().endswith(".xlsx")]
    random.Random(a.seed).shuffle(files)
    print(f"{len(files)} Excel files. Showing {min(a.n, len(files))} random ones (read-only):\n")
    for p in files[:a.n]:
        try: sheets, rows = _xlsx_peek(p)
        except Exception as e: print(f"{os.path.basename(p)}  [unreadable: {e}]\n"); continue
        print(f"FILE: {os.path.basename(p)}\n  sheets: {sheets[:4]}")
        for r in rows: print("  | " + " | ".join(r))
        print()

_AR = re.compile(r"[\u0600-\u06FF]")
_GENERIC = re.compile(r"^(sheet\d*|page\b.*|part\b.*|new balance.*|item|description|unit|qty|quantity|amount|total|sr\.?|no\.?|date|ref.*|remarks?|s\.?n\.?|pay-?\d*|[\d\s.,/()-]+)$", re.I)
_LABELS = [("project", ("اسم المشروع", "المشروع")), ("subject", ("الموضوع", "البيان")),
           ("contractor", ("المقاول", "اسم المقاول")), ("contract", ("رقم العقد", "العقد"))]

def _nrm(t):
    return re.sub(r"[\u064B-\u065F]", "", t).translate(_NORM).strip()

def _ar_ratio(t):
    letters = [c for c in t if c.isalpha()]
    return sum(1 for c in letters if _AR.match(c)) / len(letters) if letters else 0

def _xlsx_arabic_name(path):
    sheets, rows = _xlsx_peek(path, nrows=40, ncells=14, maxlen=90)
    text = "\n".join(" ".join(r) for r in rows)
    num, _ = _case_from_text(text)
    if num: return "قضية " + num, "case number"
    found = {}
    for r in rows:
        for i, cell in enumerate(r):
            c = _nrm(cell)
            for key, names in _LABELS:
                if key in found: continue
                for nm in names:
                    nmn = _nrm(nm)
                    if c.startswith(nmn) and len(c) <= len(nmn) + 80 and (":" in c or c.rstrip(": ") == nmn):
                        val = cell.split(":", 1)[1].strip() if ":" in cell and len(cell.split(":", 1)[1].strip()) >= 3 else ""
                        j = i + 1
                        while not val and j < len(r):
                            nxt = r[j].strip(); j += 1
                            if len(nxt) >= 3 and not nxt.endswith(":") and not _GENERIC.match(nxt): val = nxt
                        if val: found[key] = val
                        break
    parts = [found[k] for k in ("project", "subject") if k in found][:1]
    if "contractor" in found: parts.append(found["contractor"])
    if parts and any(_ar_ratio(x) > 0.5 for x in parts):
        return sanitize(" - ".join(parts), 80), "labels"
    best = ""
    for r in rows:
        for cell in r:
            c = re.sub(r"\s+", " ", cell).strip()
            if 6 <= len(c) <= 90 and _ar_ratio(c) > 0.6 and not c.endswith(":") and not _GENERIC.match(c):
                if len(c) > len(best): best = c
    if best: return sanitize(best, 80), "arabic title"
    for sh in sheets:
        if sh and _ar_ratio(sh) > 0.6 and not _GENERIC.match(sh): return sanitize(sh, 80), "sheet name"
    return None, None

def cmd_xlsxnames(a):
    """Rename Excel files in place using Arabic-first, content-based names (case number, labeled fields, title)."""
    root = os.path.abspath(a.folder)
    print("Scanning Excel files (progress every 200)...", flush=True)
    rows, taken, seen, keep, errs = [], set(), 0, 0, 0
    for p in iter_files(root):
        if not p.lower().endswith(".xlsx"): continue
        seen += 1
        if seen % 200 == 0: print(f"  checked {seen}, to rename {len(rows)}", flush=True)
        try: name, src = _xlsx_arabic_name(p)
        except Exception: errs += 1; continue
        if not name: keep += 1; continue
        stem = os.path.splitext(os.path.basename(p))[0]
        if re.fullmatch(re.escape(name) + r"(_\d+)?", stem): taken.add(p.lower()); continue
        rows.append((p, os.path.join(os.path.dirname(p), unique(os.path.dirname(p), name, ".xlsx", taken)), src))
    by = collections.Counter(r[2] for r in rows)
    print(f"\nExcel files: {seen}\nWill be renamed: {len(rows)}  {dict(by)}\nNo Arabic title found (left unchanged): {keep}   Errors: {errs}")
    if not a.apply:
        random.Random(1).shuffle(rows)
        print(f"\nPREVIEW ONLY. {min(a.n, len(rows))} examples:\n")
        for old, new, src in rows[:a.n]: print(f"{os.path.basename(old)}\n   -> {os.path.basename(new)}   [{src}]")
        if not (rows and sys.stdin.isatty()): return
        if input(f"\nType YES (capitals) to rename all {len(rows)} files now: ").strip() != "YES":
            print("Stopped. Nothing was renamed."); return
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    ok = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["old_path", "new_path", "status", "source", "note"])
        for old, new, src in rows:
            try:
                if os.path.lexists(new): raise FileExistsError(new)
                os.rename(old, new); w.writerow([old, new, "renamed", src, ""]); ok += 1
            except Exception as e:
                w.writerow([old, "", "skipped", "", str(e)])
            lf.flush()
    print(f"Renamed: {ok}\nUndo log: {logp}")

def cmd_pdfstat(a):
    """Read-only: how many PDFs hold real text versus scanned images (approximate, built-in reader)."""
    root = os.path.abspath(a.folder)
    files = [p for p in iter_files(root) if p.lower().endswith(".pdf")]
    random.Random(a.seed).shuffle(files)
    sample = files[:a.n]
    cnt = collections.Counter(); lines = []
    for p in sample:
        try:
            with open(p, "rb") as f: d = f.read(4_000_000)
            blob = d
            for m in re.finditer(rb"stream\r?\n(.*?)\r?\n?endstream", d, re.S):
                if len(m.group(1)) < 400_000:
                    try: blob += zlib.decompress(m.group(1))[:400_000]
                    except Exception: pass
            fonts = len(re.findall(rb"/Type\s*/Font\b", blob)); imgs = len(re.findall(rb"/Subtype\s*/Image", blob))
            pages = len(re.findall(rb"/Type\s*/Page\b", blob)); tounicode = b"/ToUnicode" in blob
            kind = "text" if fonts and not imgs else "scanned/image" if imgs and not fonts else "mixed" if imgs and fonts else "unknown"
            cnt[kind] += 1
            if len(lines) < 10: lines.append(f"{os.path.basename(p)[:45]:45s} pages~{pages:3d} fonts {fonts:3d} images {imgs:3d} ToUnicode {tounicode}  => {kind}")
        except Exception as e:
            cnt["unreadable"] += 1
    print(f"PDF files: {len(files)}   sampled: {len(sample)}\n  " + "\n  ".join(f"{k}: {v}" for k, v in cnt.most_common()) + "\n")
    for l in lines: print(l)

def _find_tesseract():
    import shutil as _sh
    t = _sh.which("tesseract")
    if t: return t
    for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(c): return c
    return None

_STOP_LINES = ("وزار", "دول", "بسم", "الرحم", "المملك", "جمهور", "ختم", "صور", "لعدل", "العدل", "عنايه", "المحترم", "تحيه", "السلام عليكم")

def _ocr_first_page(pdf, tess, dpi=120, top=0.5):
    try: import pymupdf
    except ImportError: import fitz as pymupdf
    import subprocess, tempfile
    doc = pymupdf.open(pdf)
    if doc.page_count == 0: return ""
    pg = doc[0]; r = pg.rect
    pix = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, clip=pymupdf.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * top))
    doc.close()
    env = dict(os.environ, OMP_THREAD_LIMIT="1")          # one thread per OCR job; parallelism is across files
    with tempfile.TemporaryDirectory() as td:
        img = os.path.join(td, "p.png"); pix.save(img)
        res = subprocess.run([tess, img, "stdout", "-l", "ara+eng", "--psm", "6"], capture_output=True, timeout=120, env=env)
    return res.stdout.decode("utf-8", "ignore")

def _ocr_pages(pdf, tess, dpi=120, top=1.0, npages=3):
    """OCR up to npages pages of a PDF (whole page height by default); returns a list of texts."""
    try: import pymupdf
    except ImportError: import fitz as pymupdf
    import subprocess, tempfile
    doc = pymupdf.open(pdf); n = doc.page_count; texts = []
    env = dict(os.environ, OMP_THREAD_LIMIT="1")
    with tempfile.TemporaryDirectory() as td:
        for i in range(min(n, npages if npages > 0 else n)):
            pg = doc[i]; r = pg.rect
            pix = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, clip=pymupdf.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * top))
            img = os.path.join(td, f"p{i}.png"); pix.save(img)
            try:
                res = subprocess.run([tess, img, "stdout", "-l", "ara+eng", "--psm", "6"], capture_output=True, timeout=120, env=env)
                texts.append(res.stdout.decode("utf-8", "ignore"))
            except Exception: texts.append("")
    doc.close()
    return texts

def _ocr_task(args):
    p, tess, dpi, top = args
    try: return p, _ocr_first_page(p, tess, dpi, top)
    except BaseException: return p, ""

def _fix_dir(text):
    """OCR/PDF text sometimes comes back in visual (reversed) order; put Arabic lines back in reading order."""
    out = []
    for ln in text.splitlines():
        toks = ln.split()
        ar = [t for t in toks if _AR.search(t)]
        if len(ar) >= 2:
            fwd = sum(1 for t in ar if _nrm(t).startswith(("ال", "وال", "بال", "لل", "في", "من", "علي", "الي")))
            rev = sum(1 for t in ar if _nrm(t).endswith(("لا", "ةيف", "نم", "ىلع")))
            if rev > fwd:
                ln = " ".join(t if re.fullmatch(r"[\d/.\-:()]+", t) else t[::-1] for t in reversed(toks))
        out.append(ln)
    return "\n".join(out)

_OCR_CASE_RE = re.compile(r"(?:القضي\w?|قضي\w?|الدعو[يه]\w?)[^\d]{0,10}?(\d{1,6})(?:\s*[/\\\-]\s*((?:19|20)?\d{2})|\s*(?:لسنه|لعام|عام|سنه)\s*((?:19|20)\d{2}))?")

_DOC_WORDS = ("مذكره", "صحيفه", "حكم", "اعلان", "محضر", "شهاده", "عقد", "تقرير", "خطاب", "كتاب", "انذار", "توكيل", "وكاله",
              "اقرار", "اخطار", "فاتوره", "كشف", "قرار", "مناقصه", "ترخيص", "تفويض", "مستخلص", "طلب", "عريضه", "استئناف",
              "تظلم", "طعن", "دعوي", "محضر", "ايصال", "شيك", "سند", "بلاغ", "شكوي", "عرض سعر", "تعهد", "اتفاقيه", "ملحق")

_FILLER = ("رقم", "تاريخ", "بتاريخ", "السيد", "السيده", "الاستاذ", "مقدم", "مقدمه", "الي", "علي", "الى", "من", "في", "ذات")
_GENERIC_TYPES = ("عقد", "كتاب", "طلب", "قرار", "سند", "حكم", "بلاغ", "ملحق", "تقرير", "خطاب", "ايصال", "شيك", "كشف")

def _clean_tok(t):
    return len(t) >= 3 and _ar_ratio(t) > 0.9

def _doc_type(text):
    """Document type phrase from the top lines: a document-type word plus up to 3 clean following words."""
    for ln in text.splitlines():
        toks = ln.split()
        nt = [_nrm(t) for t in toks]
        if any(sw in _nrm(ln) for sw in _STOP_LINES) and not any(nt_.startswith(w) for nt_ in nt for w in _DOC_WORDS): continue
        for i, t in enumerate(nt):
            if any(t == w or (t.startswith(w) and len(t) <= len(w) + 3) for w in _DOC_WORDS) and _ar_ratio(toks[i]) > 0.9 and len(t) >= 3:
                phrase = [toks[i]]
                for x, nx in zip(toks[i + 1:i + 4], nt[i + 1:i + 4]):
                    if _clean_tok(x) and nx not in _FILLER and not any(sw in nx for sw in _STOP_LINES): phrase.append(x)
                    else: break
                return sanitize(" ".join(phrase), 50), len(phrase), nt[i]
    return None, 0, ""

_SUBJ_KEYS = ("الموضوع", "موضوع", "بخصوص", "بشان", "بشأن")

def _subject(text):
    """The 'الموضوع:' (subject) line of a letter: up to 8 clean words after the keyword."""
    for ln in text.splitlines()[:40]:
        toks = ln.split(); nt = [_nrm(t).strip(":-") for t in toks]
        for i, t in enumerate(nt):
            if any(t == k or t.startswith(k) and len(t) <= len(k) + 2 for k in _SUBJ_KEYS) and _ar_ratio(toks[i]) > 0.8:
                rest = [x.strip(":-") for x in toks[i + 1:]]
                words = [x for x in rest if _clean_tok(x) and _nrm(x) not in _FILLER]
                if len(words) >= 2: return sanitize(" ".join(words[:8]), 60)
    return None

def _find_date(t):
    for m in re.finditer(r"(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})|(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})", t):
        g = m.groups()
        y, mo, d = (int(g[0]), int(g[1]), int(g[2])) if g[0] else (int(g[5]), int(g[4]), int(g[3]))
        if 1990 <= y <= 2035 and 1 <= mo <= 12 and 1 <= d <= 31: return f"{y}-{mo:02d}-{d:02d}"
    return None

def _name_from_ocr(text, use_titles):
    text = _fix_dir(text)
    t = re.sub(r"\s+", " ", re.sub(r"[\u064B-\u065F]", "", text).translate(_AR_DIGITS).translate(_NORM))
    num = None
    m = _OCR_CASE_RE.search(t)
    if m:
        yr = m.group(2) or m.group(3)
        num = m.group(1) + ("-" + yr if yr else "")
        if len(m.group(1)) < 2 and not yr: num = None
    date = _find_date(t)
    subj = _subject(text) if use_titles else None
    if subj:
        base = f"{subj} {date}" if date else subj
        return (f"{base} - قضية {num}" if num else base), "subject"
    dtype, nwords, kw = _doc_type(text) if use_titles else (None, 0, "")
    if dtype and nwords == 1 and kw in _GENERIC_TYPES and not (num or date): dtype = None   # a bare "كتاب"/"عقد" says nothing
    if dtype and num: return f"{dtype} - قضية {num}", "type + case number"
    if num: return "قضية " + num, "case number"
    if dtype and date: return f"{dtype} {date}", "type + date"
    if dtype: return dtype, "document type"
    if use_titles:
        for ln in text.splitlines():
            ln = re.sub(r"\s+", " ", ln).strip()
            words = [w for w in ln.split() if _clean_tok(w)]
            if 12 <= len(ln) <= 90 and len(words) >= 3 and len(words) >= 0.8 * len(ln.split()) and not any(sw in _nrm(ln) for sw in _STOP_LINES):
                return sanitize(" ".join(words[:8]), 60), "title line"
    return None, None

def cmd_pdfocr(a):
    """OCR page 1 of scanned PDFs and rename by case number (optionally by first Arabic title line)."""
    tess = _find_tesseract()
    if not tess: sys.exit("Tesseract OCR is not installed (needs the Arabic language data). Install it first, then run again.")
    try:
        try: import pymupdf
        except ImportError: import fitz
    except ImportError: sys.exit("PyMuPDF is not installed. Run:  py -m pip install pymupdf")
    import subprocess
    mine = os.path.join(os.path.expanduser("~"), "tessdata")           # user-writable language folder (no admin needed)
    if os.path.exists(os.path.join(mine, "ara.traineddata")):
        os.environ["TESSDATA_PREFIX"] = mine
        print(f"Using language files from {mine}", flush=True)
    try:
        out = subprocess.run([tess, "--list-langs"], capture_output=True).stdout.decode("utf-8", "ignore").splitlines()
        langs = [l.strip() for l in out[1:] if l.strip()]
    except Exception: langs = []
    if "ara" not in langs:
        sys.exit("Tesseract has NO Arabic language data (it only has: " + ", ".join(langs) + ").\n"
                 "Put ara.traineddata and eng.traineddata in the folder  " + mine + "  and run again.")
    root = os.path.abspath(a.folder)
    files = [p for p in iter_files(root) if p.lower().endswith(".pdf")]
    cache_p = os.path.join(root, "ocr_cache.csv"); cache = {}
    if a.apply and os.path.exists(cache_p):
        with open(cache_p, newline="", encoding="utf-8-sig") as f:
            for r in csv.reader(f):
                if r: cache[r[0]] = r[1] if len(r) > 1 else ""
    files = [p for p in files if p not in cache]
    random.Random(a.seed).shuffle(files) if not a.apply else None
    lim = a.limit if a.limit else (len(files) if a.all else 20)
    files = files[:lim]
    print(f"Tesseract: {tess}\nPDFs to process now: {len(files)}" + ("" if a.apply else "   (PREVIEW: nothing will be renamed)"), flush=True)
    logp = os.path.join(root, f"rename_log_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    taken, done, named, t0 = set(), 0, 0, time.time()
    lf = cf = None
    if a.apply:
        lf = open(logp, "w", newline="", encoding="utf-8-sig"); lw = csv.writer(lf); lw.writerow(["old_path", "new_path", "status", "source", "note"])
        cf = open(cache_p, "a", newline="", encoding="utf-8-sig"); cw = csv.writer(cf)
    from concurrent.futures import ProcessPoolExecutor
    workers = a.workers or max(1, (os.cpu_count() or 2) - 1)
    print(f"Using {workers} parallel workers; reading the top {int(a.top*100)}% of page 1 at {a.dpi} dpi.", flush=True)
    jobs = [(p, tess, a.dpi, a.top) for p in files]
    with ProcessPoolExecutor(max_workers=workers) as ex:
        for p, text in ex.map(_ocr_task, jobs, chunksize=2):
            done += 1
            name, src = _name_from_ocr(text, not a.no_titles)
            if not a.apply:
                snip = re.sub(r"\s+", " ", text).strip()[:90]
                print(f"{os.path.basename(p)}\n   OCR: {snip}\n   -> {name + '.pdf' + '   [' + src + ']' if name else '(kept: no case number found)'}", flush=True)
            else:
                if name:
                    try:
                        new = os.path.join(os.path.dirname(p), unique(os.path.dirname(p), name, ".pdf", taken))
                        if os.path.lexists(new): raise FileExistsError(new)
                        os.rename(p, new); lw.writerow([p, new, "renamed", src, ""]); named += 1
                    except Exception as e: lw.writerow([p, "", "skipped", "", str(e)]); cw.writerow([p, ""])
                else: cw.writerow([p, ""])
                lf.flush(); cf.flush()
                if done % 100 == 0:
                    rate = done / max(1, time.time() - t0); left = (len(files) - done) / max(rate, 1e-9)
                    print(f"  {done}/{len(files)}  renamed: {named}  about {left/60:.0f} min left", flush=True)
            if name and not a.apply: named += 1
    print(f"\nDone. Processed: {done}   {'Renamed' if a.apply else 'Would be renamed'}: {named}" + (f"\nUndo log: {logp}" if a.apply else ""))

def _flex(word):
    """Regex for an Arabic word that tolerates أ/إ/ا, ي/ى and ة/ه spelling variants."""
    m = {"ا": "[اأإآ]", "أ": "[اأإآ]", "إ": "[اأإآ]", "ي": "[يى]", "ى": "[يى]", "ه": "[هة]", "ة": "[هة]"}
    return "".join(m.get(c, re.escape(c)) for c in word)

_PLAINTIFF = "|".join(_flex(w) for w in ("المدعي", "المدعيه", "الطالب", "المستانف", "المستأنف", "المعارض", "الطاعن"))
_DEFENDANT = "|".join(_flex(w) for w in ("المدعي عليه", "المدعي عليها", "المطلوب ضده", "المستانف ضده", "المستأنف ضده", "المطعون ضده"))
_PLAINTIFF_RE = re.compile(r"(?:" + _PLAINTIFF + r")\s*[:：]\s*([^\n|:]{3,60})")
_DEFENDANT_RE = re.compile(r"(?:" + _DEFENDANT + r")\s*[:：]\s*([^\n|:]{3,60})")
_VS_RE = re.compile(r"([\u0600-\u06FF][\u0600-\u06FF ]{2,30}?)\s+ضد\s+([\u0600-\u06FF][\u0600-\u06FF ]{2,30})")
_PARTY_STOP = ("السيد", "السيده", "الاستاذ", "الأستاذ", "شركه", "شركة", "مؤسسه", "مؤسسة", "وكيله", "وكيلها", "وكيل")

def _party_name(raw):
    toks = [t for t in re.sub(r"[^\w\s]", " ", raw).split() if _clean_tok(t) and _nrm(t) not in ("السيد", "السيده", "الاستاذ")]
    return " ".join(toks[:3]) if len(toks) >= 1 and sum(len(t) for t in toks[:3]) >= 5 else ""

_BIDI = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069]")

def _party_from_line(toks, i):
    """Clean words after label token i (skipping noise and a following 'عليه')."""
    vals = []
    for x in toks[i + 1:i + 8]:
        x = x.strip(":：-.,")
        nx = _nrm(x)
        if nx.startswith("علي") and not vals: continue
        if nx in ("السيد", "السيده", "الاستاذ", "المحامي", "وكيله", "وكيلها", "وكيل"): continue
        if _clean_tok(x): vals.append(x)
        elif vals: break
        if len(vals) == 3: break
    return " ".join(vals) if sum(len(v) for v in vals) >= 5 else ""

def _extract_case_info(text):
    """(case_number, plaintiff, defendant) from OCR text of page 1; any may be ''."""
    text = _BIDI.sub("", _fix_dir(text))
    t = re.sub(r"\s+", " ", re.sub(r"[\u064B-\u065F]", "", text).translate(_AR_DIGITS).translate(_NORM))
    num = ""
    m = _OCR_CASE_RE.search(t)
    if m:
        yr = m.group(2) or m.group(3)
        num = m.group(1) + ("-" + yr if yr else "")
        if len(m.group(1)) < 2 and not yr: num = ""
    pl = df = ""
    for ln in text.splitlines():
        toks = ln.split(); nt = [_nrm(x).strip(":：") for x in toks]
        for i, x in enumerate(nt):
            if not x.startswith("المدع") and not x.startswith("المطلوب") and not x.startswith("المستان") and not x.startswith("الطالب"): continue
            nxt = nt[i + 1] if i + 1 < len(nt) else ""
            is_def = (nxt.startswith("علي") and x.startswith("المدع")) or x.startswith("المطلوب") or (x.startswith("المستان") and nxt.startswith("ضد"))
            val = _party_from_line(toks, i + (1 if is_def and nxt.startswith("علي") else 0))
            if val and is_def and not df: df = val
            elif val and not is_def and not pl: pl = val
            break
    if not (pl and df):
        v = _VS_RE.search(text)
        if v:
            pl = pl or _party_name(v.group(1)); df = df or _party_name(v.group(2))
    return num, pl, df

def _made_dirs(root):
    """Folder names (directly under root) created by earlier copy runs, from the copy logs."""
    out = set()
    try: logs = [f for f in os.listdir(root) if f.startswith("copy_log_") and f.endswith(".csv")]
    except OSError: return out
    for lg in logs:
        with open(os.path.join(root, lg), newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if r.get("copy"):
                    rel = os.path.relpath(os.path.dirname(r["copy"]), root)
                    if rel != "." and os.sep not in rel: out.add(rel)
    return out

def _pdf_case_files(root, ext=".pdf"):
    made = _made_dirs(root)
    for dp, dn, fn in os.walk(root):
        dn[:] = [d for d in dn if not d.startswith("قضية ") and d not in SYSTEM_DIRS and not (dp == root and d in made)]
        for n in sorted(fn):
            if n.lower().endswith(ext): yield os.path.join(dp, n)

_NAME_CASE = re.compile(r"^قضية (\d+(?:-\d+)?)(?:_\d+)?$")

def _case_task(args):
    p, tess, dpi, top, npages = args
    try: return p, _ocr_pages(p, tess, dpi, top, npages)
    except BaseException: return p, []

def _merge_case_info(texts):
    """Vote across pages: most frequent case number (years preferred); first parties found."""
    nums, pls, dfs = collections.Counter(), [], []
    for t in texts:
        num, pl, df = _extract_case_info(t)
        if num: nums[num] += 2 if "-" in num else 1
        if pl: pls.append(pl)
        if df: dfs.append(df)
    num = nums.most_common(1)[0][0] if nums else ""
    return num, (collections.Counter(pls).most_common(1)[0][0] if pls else ""), (collections.Counter(dfs).most_common(1)[0][0] if dfs else "")

def _write_case_summary(root, cache_p, sp_name="cases_summary.csv"):
    rows = collections.defaultdict(lambda: {"n": 0, "pl": collections.Counter(), "df": collections.Counter()})
    with open(cache_p, newline="", encoding="utf-8-sig") as f:
        for r in csv.reader(f):
            if len(r) < 4 or not r[1]: continue
            d = rows[r[1]]; d["n"] += 1
            if r[2]: d["pl"][r[2]] += 1
            if r[3]: d["df"][r[3]] += 1
    sp = os.path.join(root, sp_name)
    old = {}
    if os.path.exists(sp):                                         # keep the user's earlier edits
        with open(sp, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f): old[r["case"]] = r
    with open(sp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["case", "files", "suggested_parties", "final_name", "merge_into", "plaintiff_candidates", "defendant_candidates"])
        for case, d in sorted(rows.items(), key=lambda kv: -kv[1]["n"]):
            pl = d["pl"].most_common(1)[0][0] if d["pl"] else ""; df = d["df"].most_common(1)[0][0] if d["df"] else ""
            sug = f"{pl} ضد {df}" if pl and df else (pl or df)
            o = old.get(case, {})
            w.writerow([case, d["n"], sug, o.get("final_name", ""), o.get("merge_into", ""),
                        " / ".join(k for k, _ in d["pl"].most_common(3)), " / ".join(k for k, _ in d["df"].most_common(3))])
    return sp, len(rows)

def cmd_pdfcases(a):
    """Group scanned PDFs by case number into one folder per case (copies). Step 1 analyses; --apply copies."""
    root = os.path.abspath(a.folder)
    ext = ".docx" if a.docx else ".pdf"
    cache_p = os.path.join(root, ext[1:] + "_case_cache.csv")
    sp_name = "cases_summary.csv" if ext == ".pdf" else "cases_summary_docx.csv"
    if not a.apply and a.docx:
        done_set = set()
        if os.path.exists(cache_p):
            with open(cache_p, newline="", encoding="utf-8-sig") as f: done_set = {r[0] for r in csv.reader(f) if r}
        files = [q for q in _pdf_case_files(root, ".docx") if q not in done_set]
        if a.limit: files = files[:a.limit]
        print(f"Word files to analyse now: {len(files)}   (already done: {len(done_set)})", flush=True)
        with open(cache_p, "a", newline="", encoding="utf-8-sig") as cf:
            cw = csv.writer(cf)
            for i, q in enumerate(files, 1):
                try: text = docx_text(q)
                except Exception: text = ""
                num, pl, df = _extract_case_info(text)
                m = _NAME_CASE.match(os.path.splitext(os.path.basename(q))[0])
                if m: num = m.group(1)                       # a name already built from the case number wins
                cw.writerow([q, num, pl, df]); cf.flush()
                if i % 500 == 0: print(f"  {i}/{len(files)}", flush=True)
        sp, ncases = _write_case_summary(root, cache_p, sp_name)
        print(f"\nDone. Cases found: {ncases}\nOpen this file in Excel and fill column final_name (e.g. أحمد ضد محمد), then save it as CSV UTF-8:\n  {sp}\nThen run the same command with --apply to copy the files.")
        return
    if not a.apply:
        tess = _find_tesseract()
        if not tess: sys.exit("Tesseract OCR is not installed.")
        mine = os.path.join(os.path.expanduser("~"), "tessdata")
        if os.path.exists(os.path.join(mine, "ara.traineddata")): os.environ["TESSDATA_PREFIX"] = mine
        done_set = set()
        if os.path.exists(cache_p):
            with open(cache_p, newline="", encoding="utf-8-sig") as f: done_set = {r[0] for r in csv.reader(f) if r}
        if a.rescan and os.path.exists(cache_p):
            os.replace(cache_p, cache_p + ".bak"); done_set = set(); print("Old results kept as pdf_case_cache.csv.bak; analysing again.", flush=True)
        files = [p for p in _pdf_case_files(root) if p not in done_set]
        if a.limit: files = files[:a.limit]
        print(f"PDFs to analyse now: {len(files)}   (already done: {len(done_set)})   pages per file: {'all' if a.pages == 0 else a.pages}   page height read: {int(a.top*100)}%", flush=True)
        from concurrent.futures import ProcessPoolExecutor
        workers = a.workers or max(1, (os.cpu_count() or 2) - 1)
        t0, n = time.time(), 0
        with open(cache_p, "a", newline="", encoding="utf-8-sig") as cf, ProcessPoolExecutor(max_workers=workers) as ex:
            cw = csv.writer(cf)
            for p, texts in ex.map(_case_task, [(p, tess, a.dpi, a.top, a.pages) for p in files], chunksize=2):
                n += 1
                num, pl, df = _merge_case_info(texts)
                cw.writerow([p, num, pl, df]); cf.flush()
                if n % (100 if len(files) >= 1000 else 10) == 0:
                    rate = n / max(1, time.time() - t0)
                    print(f"  {n}/{len(files)}  {1/max(rate,1e-9):.1f} s per file  about {(len(files) - n) / max(rate, 1e-9) / 60:.0f} min left", flush=True)
        sp, ncases = _write_case_summary(root, cache_p, sp_name)
        print(f"\nDone. Cases found: {ncases}\nOpen this file in Excel. Type the parties in column final_name (e.g. أحمد ضد محمد) or leave it empty; merge_into = another case number to merge into:\n  {sp}\nThen run the same command with --apply to copy the files.")
        return
    sp = os.path.join(root, sp_name)
    if not (os.path.exists(sp) and os.path.exists(cache_p)): sys.exit("Run the analysis step first (without --apply).")
    folder_of = {}
    with open(sp, newline="", encoding="utf-8-sig") as f:
        srows = list(csv.DictReader(f))
    for r in srows:
        base = "قضية " + r["case"]
        folder_of[r["case"]] = sanitize(base + (" - " + r["final_name"].strip() if r["final_name"].strip() else ""), 100)
    for r in srows:
        if r["merge_into"].strip() and r["merge_into"].strip() in folder_of: folder_of[r["case"]] = folder_of[r["merge_into"].strip()]
    jobs = []
    with open(cache_p, newline="", encoding="utf-8-sig") as f:
        for r in csv.reader(f):
            if len(r) >= 2 and r[1] in folder_of and os.path.exists(r[0]): jobs.append((r[0], folder_of[r[1]]))
    need = sum(os.path.getsize(src) for src, _ in jobs)
    free = shutil.disk_usage(root).free
    print(f"Files to copy: {len(jobs)}   size: {need/1e9:.2f} GB   free space: {free/1e9:.2f} GB   case folders: {len(set(d for _, d in jobs))}", flush=True)
    if need * 1.02 > free and not a.move: sys.exit("Not enough free space to COPY. Free up space, or add --move (moves instead of copying).")
    if not a.yes:
        if not sys.stdin.isatty() or input("Type YES (capitals) to " + ("move" if a.move else "copy") + " now: ").strip() != "YES":
            print("Stopped. Nothing was copied."); return
    logp = os.path.join(root, f"copy_log_{time.strftime('%Y%m%d_%H%M%S')}.csv"); ok = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["source", "copy", "mode"])
        for src, folder in jobs:
            try:
                dd = os.path.join(root, folder); os.makedirs(dd, exist_ok=True)
                name = os.path.basename(src); dst = os.path.join(dd, name); k = 0
                while os.path.lexists(dst):
                    k += 1; dst = os.path.join(dd, f"{os.path.splitext(name)[0]}_{k}{os.path.splitext(name)[1]}")
                (shutil.move if a.move else shutil.copy2)(src, dst)
                w.writerow([src, dst, "move" if a.move else "copy"]); ok += 1
            except Exception as e: w.writerow([src, "", f"skipped: {e}"])
            lf.flush()
    print(f"Done: {ok} files {'moved' if a.move else 'copied'}.\nLog: {logp}\nTo undo: py rr.py pdfcases-undo \"{root}\"")

def cmd_pdfcases_undo(a):
    root = os.path.abspath(a.folder)
    logs = sorted(f for f in os.listdir(root) if f.startswith("copy_log_") and f.endswith(".csv"))
    if not logs: sys.exit("No copy_log found.")
    n = 0
    with open(os.path.join(root, logs[-1]), newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["mode"] == "copy" and r["copy"] and os.path.exists(r["copy"]) and os.path.exists(r["source"]):
                os.remove(r["copy"]); n += 1
            elif r["mode"] == "move" and r["copy"] and os.path.exists(r["copy"]) and not os.path.exists(r["source"]):
                os.rename(r["copy"], r["source"]); n += 1
    for d in _made_dirs(root) | {d for d in os.listdir(root) if d.startswith("قضية ")}:
        dd = os.path.join(root, d)
        if os.path.isdir(dd) and not os.listdir(dd): os.rmdir(dd)
    print(f"Undone: {n} files (empty case folders removed).")

def _party_task(args):
    """OCR pages one by one and stop as soon as a plaintiff/defendant name is found."""
    p, tess, dpi, top, maxpages = args
    try:
        try: import pymupdf
        except ImportError: import fitz as pymupdf
        import subprocess, tempfile
        doc = pymupdf.open(p); n = doc.page_count
        env = dict(os.environ, OMP_THREAD_LIMIT="1"); pl = df = ""; read = 0
        with tempfile.TemporaryDirectory() as td:
            for i in range(min(n, maxpages if maxpages > 0 else n)):
                pg = doc[i]; r = pg.rect
                pix = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, clip=pymupdf.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * top))
                img = os.path.join(td, f"p{i}.png"); pix.save(img); read += 1
                try: txt = subprocess.run([tess, img, "stdout", "-l", "ara+eng", "--psm", "6"], capture_output=True, timeout=120, env=env).stdout.decode("utf-8", "ignore")
                except Exception: txt = ""
                _, a_, b_ = _extract_case_info(txt)
                pl, df = pl or a_, df or b_
                if pl or df: break
        doc.close()
        return p, pl, df, read
    except BaseException:
        return p, "", "", 0

def _party_label(pl, df):
    return f"{pl} ضد {df}" if pl and df else (pl or df)

def _cluster_parties(labels):
    """Greedy fuzzy grouping of near-identical party labels. Returns {label: group_label}."""
    import difflib
    cnt = collections.Counter(labels); groups = []        # [(key, group_label)]
    out = {}
    for lab, _ in cnt.most_common():
        key = _nrm(lab).replace(" ", "")
        for gk, gl in groups:
            if abs(len(gk) - len(key)) <= 4 and difflib.SequenceMatcher(None, gk, key).ratio() >= 0.82:
                out[lab] = gl; break
        else:
            groups.append((key, lab)); out[lab] = lab
    return out

def _write_party_summary(root, cache_p, sp_name="parties_summary.csv"):
    items = []
    with open(cache_p, newline="", encoding="utf-8-sig") as f:
        for r in csv.reader(f):
            if len(r) >= 3 and (r[1] or r[2]): items.append((r[0], _party_label(r[1], r[2])))
    mapping = _cluster_parties([l for _, l in items])
    groups = collections.defaultdict(lambda: {"n": 0, "variants": collections.Counter()})
    for _, l in items: g = groups[mapping[l]]; g["n"] += 1; g["variants"][l] += 1
    sp = os.path.join(root, sp_name); old = {}
    if os.path.exists(sp):
        with open(sp, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f): old[r["group"]] = r
    with open(sp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["group", "files", "final_name", "merge_into", "variants_seen"])
        for g, d in sorted(groups.items(), key=lambda kv: -kv[1]["n"]):
            o = old.get(g, {})
            w.writerow([g, d["n"], o.get("final_name", ""), o.get("merge_into", ""), " / ".join(k for k, _ in d["variants"].most_common(4))])
    return sp, len(groups), mapping

def _do_copy(root, jobs, a):
    need = sum(os.path.getsize(src) for src, _ in jobs); free = shutil.disk_usage(root).free
    print(f"Files to copy: {len(jobs)}   size: {need/1e9:.2f} GB   free space: {free/1e9:.2f} GB   folders: {len(set(d for _, d in jobs))}", flush=True)
    if need * 1.02 > free and not a.move: sys.exit("Not enough free space to COPY. Free up space, or add --move.")
    if not a.yes:
        if not sys.stdin.isatty() or input("Type YES (capitals) to " + ("move" if a.move else "copy") + " now: ").strip() != "YES":
            print("Stopped. Nothing was copied."); return
    logp = os.path.join(root, f"copy_log_{time.strftime('%Y%m%d_%H%M%S')}.csv"); ok = 0
    with open(logp, "w", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf); w.writerow(["source", "copy", "mode"])
        for src, folder in jobs:
            try:
                dd = os.path.join(root, folder); os.makedirs(dd, exist_ok=True)
                name = os.path.basename(src); dst = os.path.join(dd, name); k = 0
                while os.path.lexists(dst):
                    k += 1; dst = os.path.join(dd, f"{os.path.splitext(name)[0]}_{k}{os.path.splitext(name)[1]}")
                (shutil.move if a.move else shutil.copy2)(src, dst)
                w.writerow([src, dst, "move" if a.move else "copy"]); ok += 1
            except Exception as e: w.writerow([src, "", f"skipped: {e}"])
            lf.flush()
    print(f"Done: {ok} files {'moved' if a.move else 'copied'}.\nLog: {logp}")

def cmd_pdfparties(a):
    """Folder per dispute, named by the parties (e.g. 'أحمد ضد محمد'); reads pages until a party name is found."""
    root = os.path.abspath(a.folder)
    cache_p = os.path.join(root, "pdf_party_cache.csv")
    if not a.apply:
        tess = _find_tesseract()
        if not tess: sys.exit("Tesseract OCR is not installed.")
        mine = os.path.join(os.path.expanduser("~"), "tessdata")
        if os.path.exists(os.path.join(mine, "ara.traineddata")): os.environ["TESSDATA_PREFIX"] = mine
        done_set = set()
        if a.rescan and os.path.exists(cache_p): os.replace(cache_p, cache_p + ".bak"); print("Old results kept as .bak; analysing again.", flush=True)
        if os.path.exists(cache_p):
            with open(cache_p, newline="", encoding="utf-8-sig") as f: done_set = {r[0] for r in csv.reader(f) if r}
        files = [q for q in _pdf_case_files(root) if q not in done_set]
        if a.limit: files = files[:a.limit]
        print(f"PDFs to analyse now: {len(files)}   (already done: {len(done_set)})   up to {a.maxpages or 'all'} pages each, stops when a party name is found", flush=True)
        from concurrent.futures import ProcessPoolExecutor
        workers = a.workers or max(1, (os.cpu_count() or 2) - 1); t0 = time.time(); n = found = 0
        with open(cache_p, "a", newline="", encoding="utf-8-sig") as cf, ProcessPoolExecutor(max_workers=workers) as ex:
            cw = csv.writer(cf)
            for q, pl, df, read in ex.map(_party_task, [(q, tess, a.dpi, a.top, a.maxpages) for q in files], chunksize=2):
                n += 1; found += 1 if (pl or df) else 0
                cw.writerow([q, pl, df, read]); cf.flush()
                if n % (100 if len(files) >= 1000 else 10) == 0:
                    rate = n / max(1, time.time() - t0)
                    print(f"  {n}/{len(files)}  parties found in {found}  {1/max(rate,1e-9):.1f} s per file  about {(len(files)-n)/max(rate,1e-9)/60:.0f} min left", flush=True)
        sp, ng, _ = _write_party_summary(root, cache_p)
        print(f"\nDone. Disputes (groups) found: {ng}\nOpen in Excel, fix names in final_name, merge duplicates with merge_into, save as CSV UTF-8:\n  {sp}\nThen run again with --apply.")
        return
    sp = os.path.join(root, "parties_summary.csv")
    if not (os.path.exists(sp) and os.path.exists(cache_p)): sys.exit("Run the analysis step first (without --apply).")
    _, _, mapping = _write_party_summary(root, cache_p)
    with open(sp, newline="", encoding="utf-8-sig") as f: srows = list(csv.DictReader(f))
    folder_of = {r["group"]: sanitize((r["final_name"].strip() or r["group"]), 100) for r in srows}
    for r in srows:
        m_ = r["merge_into"].strip()
        if m_ and m_ in folder_of: folder_of[r["group"]] = folder_of[m_]
    jobs = []
    with open(cache_p, newline="", encoding="utf-8-sig") as f:
        for r in csv.reader(f):
            if len(r) >= 3 and (r[1] or r[2]) and os.path.exists(r[0]):
                g = mapping.get(_party_label(r[1], r[2])); fo = folder_of.get(g)
                if fo: jobs.append((r[0], fo))
    _do_copy(root, jobs, a)

def cmd_undo(a):
    n = 0
    if os.path.isdir(a.log):                       # a folder: use its newest rename_log_*.csv
        logs = sorted(f for f in os.listdir(a.log) if f.startswith("rename_log_") and f.endswith(".csv"))
        if not logs: sys.exit(f"No rename_log_*.csv found in {a.log}")
        a.log = os.path.join(a.log, logs[-1]); print(f"Using log: {a.log}")
    with open(a.log, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            seen_rows = locals().get("seen_rows", 0) + 1
            if seen_rows % 500 == 0: print(f"  checked {seen_rows} log rows, restored {n}", flush=True)
            if r["status"] != "renamed": continue
            if a.ext and not r["new_path"].lower().endswith(a.ext.lower()): continue
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
    p = sp.add_parser("thumbs"); p.add_argument("previews"); p.add_argument("docx_root")
    p.add_argument("--all", action="store_true"); p.add_argument("--limit", type=int, default=3)
    p.add_argument("--skip-convert", action="store_true"); p.add_argument("--from-pictures", action="store_true")
    p.add_argument("--force", action="store_true"); p.set_defaults(fn=cmd_thumbs)
    p = sp.add_parser("content"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("-n", type=int, default=20); p.add_argument("--prefix", default="قضية "); p.set_defaults(fn=cmd_content)
    p = sp.add_parser("xlsxpeek"); p.add_argument("folder"); p.add_argument("-n", type=int, default=12)
    p.add_argument("--seed", type=int, default=1); p.set_defaults(fn=cmd_xlsxpeek)
    p = sp.add_parser("xlsxnames"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("-n", type=int, default=25); p.set_defaults(fn=cmd_xlsxnames)
    p = sp.add_parser("pdfstat"); p.add_argument("folder"); p.add_argument("-n", type=int, default=200)
    p.add_argument("--seed", type=int, default=1); p.set_defaults(fn=cmd_pdfstat)
    p = sp.add_parser("pdfocr"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("--all", action="store_true"); p.add_argument("--limit", type=int, default=0)
    p.add_argument("--no-titles", action="store_true", help="only rename when a case number is found")
    p.add_argument("--seed", type=int, default=1)
    p.add_argument("--workers", type=int, default=0); p.add_argument("--dpi", type=int, default=120)
    p.add_argument("--top", type=float, default=0.5, help="fraction of the page height to read, from the top"); p.set_defaults(fn=cmd_pdfocr)
    p = sp.add_parser("pdfcases"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("--move", action="store_true"); p.add_argument("--yes", action="store_true"); p.add_argument("--docx", action="store_true")
    p.add_argument("--limit", type=int, default=0); p.add_argument("--workers", type=int, default=0)
    p.add_argument("--dpi", type=int, default=120); p.add_argument("--top", type=float, default=0.6)
    p.add_argument("--pages", type=int, default=1, help="pages to read per PDF (0 = all)"); p.add_argument("--rescan", action="store_true"); p.set_defaults(fn=cmd_pdfcases)
    p = sp.add_parser("pdfparties"); p.add_argument("folder"); p.add_argument("--apply", action="store_true")
    p.add_argument("--move", action="store_true"); p.add_argument("--yes", action="store_true"); p.add_argument("--rescan", action="store_true")
    p.add_argument("--limit", type=int, default=0); p.add_argument("--workers", type=int, default=0)
    p.add_argument("--dpi", type=int, default=120); p.add_argument("--top", type=float, default=1.0)
    p.add_argument("--maxpages", type=int, default=5, help="max pages to read per PDF (0 = all)"); p.set_defaults(fn=cmd_pdfparties)
    p = sp.add_parser("pdfcases-undo"); p.add_argument("folder"); p.set_defaults(fn=cmd_pdfcases_undo)
    p = sp.add_parser("undo"); p.add_argument("log"); p.add_argument("--ext", default="", help="only undo files with this extension, e.g. .pdf"); p.set_defaults(fn=cmd_undo)
    a = ap.parse_args(); a.fn(a)

if __name__ == "__main__":
    main()
