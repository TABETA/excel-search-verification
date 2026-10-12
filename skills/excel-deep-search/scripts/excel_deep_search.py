#!/usr/bin/env python3
"""Excel deep search: find a keyword everywhere inside .xlsx/.xlsm/.xls files.

Searches cells (incl. hidden sheets, rich text, formula cached values), shapes/text boxes
(incl. groups, split runs), charts, comments, header/footer, data validation, defined
names, sheet names, document properties, image alt text. Extracts embedded images for
visual reading. Reports encrypted files. Read-only on the source files.

usage: python3 excel_deep_search.py FOLDER KEYWORD [KEYWORD ...] --out OUTDIR [--exclude NAME ...]
stdlib only; Pillow optional (contact sheets); LibreOffice optional (.xls shapes/comments).
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile, unicodedata, zipfile
import xml.etree.ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".webp")


def ln(tag):
    return tag.rsplit("}", 1)[-1]


def norm(s):
    """Loose form for 表記ゆれ: NFKC, hiragana->katakana, drop all whitespace."""
    s = unicodedata.normalize("NFKC", s)
    s = "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in s)
    return re.sub(r"\s+", "", s)


def para_text(el):
    """Text of a DrawingML/SpreadsheetML element: runs joined, paragraphs separated by \\n."""
    paras = [p for p in el.iter() if ln(p.tag) == "p" and p.tag.startswith("{" + NS_A)]
    if paras:
        out = []
        for p in paras:
            out.append("".join(t.text or "" for t in p.iter() if ln(t.tag) == "t"))
        return "\n".join(out)
    return "".join(t.text or "" for t in el.iter() if ln(t.tag) == "t")


class Book:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())

    def xml(self, name):
        try:
            return ET.fromstring(self.z.read(name))
        except Exception:
            return None

    def rels(self, part):
        d, b = os.path.split(part)
        rp = f"{d}/_rels/{b}.rels" if d else f"_rels/{b}.rels"
        out = {}
        root = self.xml(rp) if rp in self.names else None
        if root is None:
            return out
        for r in root:
            tgt = r.get("Target", "")
            if r.get("TargetMode") == "External":
                continue
            full = os.path.normpath(os.path.join(d, tgt)).replace("\\", "/") if not tgt.startswith("/") else tgt[1:]
            out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], full)
        return out


def scan_xlsx(path, emit, images, img_dir, label):
    bk = Book(path)
    covered = set()
    # shared strings
    sst = []
    if "xl/sharedStrings.xml" in bk.names:
        covered.add("xl/sharedStrings.xml")
        for si in bk.xml("xl/sharedStrings.xml"):
            sst.append("".join(t.text or "" for t in si.iter() if ln(t.tag) == "t"))
    wb = bk.xml("xl/workbook.xml"); covered.add("xl/workbook.xml")
    wrels = bk.rels("xl/workbook.xml")
    sheets = []
    for el in wb.iter():
        if ln(el.tag) == "sheet":
            rid = el.get(f"{{{NS_R}}}id")
            state = el.get("state", "visible")
            sheets.append((el.get("name"), state, wrels.get(rid, ("", ""))[1]))
            emit(f"シート名（{state}）", el.get("name"))
        elif ln(el.tag) == "definedName":
            emit(f"名前定義「{el.get('name')}」", el.text or "")
    for sname, state, spath in sheets:
        tag = f"シート「{sname}」" + ("" if state == "visible" else f"（{state}）")
        root = bk.xml(spath) if spath in bk.names else None
        covered.add(spath)
        if root is None:
            continue
        kind = ln(root.tag)
        if kind == "worksheet":
            for c in root.iter(f"{{{NS_MAIN}}}c"):
                ref, t = c.get("r"), c.get("t")
                v = c.find(f"{{{NS_MAIN}}}v"); f = c.find(f"{{{NS_MAIN}}}f")
                if t == "s" and v is not None:
                    try: emit(f"{tag} セル{ref}", sst[int(v.text)])
                    except Exception: pass
                elif t == "inlineStr":
                    emit(f"{tag} セル{ref}", "".join(x.text or "" for x in c.iter() if ln(x.tag) == "t"))
                elif v is not None and v.text:
                    emit(f"{tag} セル{ref}" + ("（数式の結果）" if f is not None else ""), v.text)
                if f is not None and f.text:
                    emit(f"{tag} セル{ref} の数式", f.text)
            for el in root.iter():
                n = ln(el.tag)
                if n in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter") and el.text:
                    emit(f"{tag} ヘッダー/フッター({n})", el.text)
                elif n == "dataValidation":
                    for fm in el:
                        if fm.text: emit(f"{tag} 入力規則 {el.get('sqref')}", fm.text)
        srels = bk.rels(spath)
        for rtype, target in srels.values():
            if target not in bk.names: continue
            covered.add(target)
            if rtype == "comments":
                for cm in bk.xml(target).iter(f"{{{NS_MAIN}}}comment"):
                    emit(f"{tag} セル{cm.get('ref')} のコメント/メモ", para_text(cm))
            elif rtype == "threadedComment":
                for cm in bk.xml(target).iter():
                    if ln(cm.tag) == "threadedComment":
                        emit(f"{tag} セル{cm.get('ref')} のスレッドコメント", "".join(x.text or "" for x in cm.iter() if ln(x.tag) == "text"))
            elif rtype == "drawing":
                scan_drawing(bk, target, tag, emit, images, img_dir, label, covered)
    # doc properties
    for p in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
        if p in bk.names:
            covered.add(p)
            for el in bk.xml(p).iter():
                if el.text and el.text.strip() and len(el) == 0:
                    emit(f"ファイルのプロパティ（{p.split('/')[-1]}:{ln(el.tag)}）", el.text)
    # safety net: any other XML part, whole text
    for n in sorted(bk.names - covered):
        if n.endswith((".xml", ".vml")) and not n.startswith("[Content_Types]"):
            root = bk.xml(n)
            if root is not None:
                txt = "".join(t.text or "" for t in root.iter() if t.text)
                emit(f"その他パーツ {n}", txt)


def scan_drawing(bk, dpath, tag, emit, images, img_dir, label, covered):
    root = bk.xml(dpath)
    drels = bk.rels(dpath)
    for el in root.iter():
        n = ln(el.tag)
        if n in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            nv = next((x for x in el.iter() if ln(x.tag) == "cNvPr"), None)
            name = nv.get("name", "") if nv is not None else ""
            descr = (nv.get("descr", "") + " " + nv.get("title", "")).strip() if nv is not None else ""
            if descr and n != "grpSp":
                emit(f"{tag} 図形「{name}」の代替テキスト", descr)
            if n == "sp":
                body = next((x for x in el if ln(x.tag) == "txBody"), None)
                if body is not None:
                    emit(f"{tag} 図形「{name}」", para_text(body))
            elif n == "pic":
                blip = next((x for x in el.iter() if ln(x.tag) == "blip"), None)
                rid = blip.get(f"{{{NS_R}}}embed") if blip is not None else None
                if rid in drels:
                    images.append(extract_image(bk, drels[rid][1], img_dir, label, f"{tag} 画像「{name}」", descr))
            elif n == "graphicFrame":
                for x in el.iter():
                    if ln(x.tag) == "chart":
                        rid = x.get(f"{{{NS_R}}}id")
                        if rid in drels and drels[rid][1] in bk.names:
                            cpath = drels[rid][1]; covered.add(cpath)
                            croot = bk.xml(cpath)
                            for tx in croot.iter():
                                if ln(tx.tag) in ("tx", "title"):
                                    t = para_text(tx) or "".join(v.text or "" for v in tx.iter() if ln(v.tag) == "v")
                                    if t: emit(f"{tag} グラフ「{name}」", t)
    # images referenced but not placed as pic (e.g. backgrounds)
    for rtype, tgt in drels.values():
        if rtype == "image" and not any(i["part"] == tgt and i["file"] == label for i in images):
            images.append(extract_image(bk, tgt, img_dir, label, f"{tag} 画像（配置不明）", ""))


def extract_image(bk, part, img_dir, label, where, descr):
    os.makedirs(img_dir, exist_ok=True)
    idx = len(os.listdir(img_dir)) + 1
    out = os.path.join(img_dir, f"{idx:04d}{os.path.splitext(part)[1].lower()}")
    with open(out, "wb") as fh:
        fh.write(bk.z.read(part))
    return {"id": idx, "file": label, "where": where, "alt": descr, "part": part, "path": out}


def is_encrypted_ole(data):
    return data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" and "EncryptedPackage".encode("utf-16-le") in data


def xls_to_xlsx(path, tmp):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    src = os.path.join(tmp, "in" + os.path.splitext(path)[1]); shutil.copy(path, src)
    subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", tmp, src],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    out = os.path.join(tmp, "in.xlsx")
    return out if os.path.exists(out) else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder"); ap.add_argument("keywords", nargs="+")
    ap.add_argument("--out", required=True); ap.add_argument("--exclude", nargs="*", default=[])
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    img_root = os.path.join(a.out, "images")
    shutil.rmtree(img_root, ignore_errors=True)
    kws = [(k, norm(k)) for k in a.keywords]
    hits, near, locked, errors, images, scanned = [], [], [], [], [], 0
    for dp, dns, fns in os.walk(a.folder):
        dns[:] = [d for d in dns if d not in a.exclude and not d.startswith(".")]
        for fn in sorted(fns):
            if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")):
                continue
            path = os.path.join(dp, fn); label = os.path.relpath(path, a.folder); scanned += 1
            seen = set()

            def emit(where, text):
                if not text: return
                for k, nk in kws:
                    if k in text:
                        key = (where, k)
                        if key not in seen:
                            seen.add(key); hits.append({"file": label, "keyword": k, "where": where, "text": text[:300]})
                    elif nk and nk in norm(text):
                        near.append({"file": label, "keyword": k, "where": where, "text": text[:300]})
            try:
                with open(path, "rb") as fh: head = fh.read()
                if is_encrypted_ole(head):
                    locked.append({"file": label, "reason": "パスワードで暗号化されている"}); continue
                if head[:2] == b"PK":
                    scan_xlsx(path, emit, images, img_root, label)
                else:  # legacy .xls
                    with tempfile.TemporaryDirectory() as tmp:
                        conv = xls_to_xlsx(path, tmp)
                        if conv:
                            scan_xlsx(conv, emit, images, img_root, label)
                        else:
                            errors.append({"file": label, "reason": "LibreOfficeが無いため.xlsは生バイト検索のみ"})
                            for k, _ in kws:
                                for enc in ("utf-16-le", "cp932"):
                                    if k.encode(enc) in head:
                                        emit(f"バイナリ一致({enc})", k)
            except Exception as e:
                errors.append({"file": label, "reason": f"{type(e).__name__}: {e}"})
    # contact sheets for visual reading
    sheets = []
    try:
        from PIL import Image, ImageDraw
        tiles = []
        for im in images:
            try:
                img = Image.open(im["path"]).convert("RGB"); img.thumbnail((600, 450)); tiles.append((im["id"], img))
            except Exception:
                im["note"] = "Pillowで開けない形式（EMF等）。個別に確認"
        for s in range(0, len(tiles), 6):
            grp = tiles[s:s + 6]
            cs = Image.new("RGB", (1240, 960), "white"); d = ImageDraw.Draw(cs)
            for k, (iid, img) in enumerate(grp):
                x, y = (k % 2) * 620 + 10, (k // 2) * 320 + 10
                img2 = img.copy(); img2.thumbnail((600, 290)); cs.paste(img2, (x, y + 20))
                d.rectangle([x, y, x + 60, y + 18], fill="black"); d.text((x + 4, y + 3), f"#{iid}", fill="white")
            p = os.path.join(a.out, f"contact_{s // 6 + 1:02d}.png"); cs.save(p); sheets.append(p)
    except ImportError:
        pass
    res = {"scanned_files": scanned, "hits": hits, "near_variants": near, "locked": locked,
           "errors": errors, "images": images, "contact_sheets": sheets}
    with open(os.path.join(a.out, "result.json"), "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1)
    files_hit = sorted({h["file"] for h in hits})
    print(f"scanned={scanned} files_with_hits={len(files_hit)} hits={len(hits)} near={len(near)} "
          f"locked={len(locked)} errors={len(errors)} images={len(images)} contact_sheets={len(sheets)}")
    print("result:", os.path.join(a.out, "result.json"))


if __name__ == "__main__":
    main()
