#!/usr/bin/env python3
"""Convert a folder of Excel files into a searchable Markdown knowledge base (read-only on sources).

  build   SRC KB [--exclude NAME ...]   create/update KB (only changed files are re-parsed)
  pending KB                            list images still waiting for transcription
  set-image KB DOC_ID IMG_NO TEXT       store the transcription of one image ("-" for no text)

KB layout:
  KB/index.md              list of documents, sheets, counts, unreadable files, pending images
  KB/docs/<doc_id>.md      one Markdown per Excel file (cells, shapes, flow arrows, comments, charts, images...)
  KB/tables/<doc_id>/<sheet>.csv   full sheet grid for calculations (row 1 = Excel row 1)
  KB/images/<doc_id>/imgN.*        embedded images
  KB/model/<doc_id>.json   parsed model (keeps image transcriptions across rebuilds)
stdlib only; LibreOffice optional (.xls).
"""
import csv, datetime as dt, hashlib, json, os, re, shutil, subprocess, sys, tempfile, zipfile
import xml.etree.ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EXCEL_EXT = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")
PENDING = "【未記入：画像を目視して set-image で記入すること】"


def ln(t):
    return t.rsplit("}", 1)[-1]


def col_letter(n):  # 0-based
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not m:
        return None
    c = 0
    for ch in m.group(1):
        c = c * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, c - 1


def safe(s):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_") or "x"


def para_text(el):
    paras = [p for p in el.iter() if p.tag == f"{{{NS_A}}}p"]
    if paras:
        return "\n".join("".join(t.text or "" for t in p.iter() if ln(t.tag) == "t") for p in paras).strip()
    return "".join(t.text or "" for t in el.iter() if ln(t.tag) == "t").strip()


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
        for r in (root if root is not None else []):
            if r.get("TargetMode") == "External":
                continue
            t = r.get("Target", "")
            full = t[1:] if t.startswith("/") else os.path.normpath(os.path.join(d, t)).replace("\\", "/")
            out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], full)
        return out


def date_styles(bk):
    """Indexes of cellXfs that are date formats."""
    st = bk.xml("xl/styles.xml") if "xl/styles.xml" in bk.names else None
    if st is None:
        return set()
    custom = {}
    for nf in st.iter(f"{{{NS_MAIN}}}numFmt"):
        custom[int(nf.get("numFmtId"))] = nf.get("formatCode", "")
    out = set()
    xfs = st.find(f"{{{NS_MAIN}}}cellXfs")
    for i, xf in enumerate(xfs if xfs is not None else []):
        fid = int(xf.get("numFmtId", 0))
        if 14 <= fid <= 22 or 45 <= fid <= 47:
            out.add(i)
        elif fid in custom:
            code = re.sub(r'"[^"]*"|\[[^\]]*\]', "", custom[fid]).lower()
            if re.search(r"[ymd]", code) and not re.fullmatch(r"[#0,.\s]*", code):
                out.add(i)
    return out


def fmt_num(v, is_date):
    try:
        f = float(v)
    except ValueError:
        return v
    if is_date and 0 < f < 2958466:
        d = dt.datetime(1899, 12, 30) + dt.timedelta(days=f)
        return d.strftime("%Y/%m/%d") if d.time() == dt.time() else d.strftime("%Y/%m/%d %H:%M")
    return str(int(f)) if f == int(f) and abs(f) < 1e15 else repr(f)


def parse_xlsx(path, doc_id, kb, note=""):
    bk = Book(path)
    sst = []
    if "xl/sharedStrings.xml" in bk.names:
        for si in bk.xml("xl/sharedStrings.xml"):
            sst.append("".join(t.text or "" for t in si.iter() if ln(t.tag) == "t"))
    dstyles = date_styles(bk)
    wb = bk.xml("xl/workbook.xml")
    wrels = bk.rels("xl/workbook.xml")
    model = {"doc_id": doc_id, "note": note, "sheets": [], "defined_names": [], "properties": {}, "images": []}
    sheet_names = {}
    for el in wb.iter():
        if ln(el.tag) == "sheet":
            sheet_names[el.get("sheetId")] = el.get("name")
            model["sheets"].append({"name": el.get("name"), "state": el.get("state", "visible"),
                                    "path": wrels.get(el.get(f"{{{NS_R}}}id"), ("", ""))[1]})
        elif ln(el.tag) == "definedName" and not (el.get("name", "").startswith("_xlnm")):
            model["defined_names"].append({"name": el.get("name"), "value": el.text or ""})
    img_dir = os.path.join(kb, "images", doc_id)
    shutil.rmtree(img_dir, ignore_errors=True)
    for sh in model["sheets"]:
        root = bk.xml(sh["path"]) if sh["path"] in bk.names else None
        sh.update(cells={}, merged=[], hidden_rows=[], hidden_cols=[], comments=[], shapes=[], edges=[],
                  charts=[], header_footer=[], validations=[], kind="worksheet")
        if root is None:
            continue
        sh["kind"] = ln(root.tag)
        for row in root.iter(f"{{{NS_MAIN}}}row"):
            if row.get("hidden") in ("1", "true"):
                sh["hidden_rows"].append(int(row.get("r")))
        for c in root.iter(f"{{{NS_MAIN}}}col"):
            if c.get("hidden") in ("1", "true"):
                sh["hidden_cols"].append(f"{col_letter(int(c.get('min'))-1)}:{col_letter(int(c.get('max'))-1)}")
        for c in root.iter(f"{{{NS_MAIN}}}c"):
            pos = parse_ref(c.get("r"))
            if not pos:
                continue
            t = c.get("t"); v = c.find(f"{{{NS_MAIN}}}v")
            if t == "s" and v is not None:
                val = sst[int(v.text)] if int(v.text) < len(sst) else ""
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter() if ln(x.tag) == "t")
            elif t == "b" and v is not None:
                val = "TRUE" if v.text == "1" else "FALSE"
            elif v is not None and v.text is not None:
                val = v.text if t in ("str", "e") else fmt_num(v.text, int(c.get("s", 0)) in dstyles)
            else:
                continue
            if val != "":
                sh["cells"][f"{pos[0]},{pos[1]}"] = val
        for el in root.iter():
            n = ln(el.tag)
            if n == "mergeCell":
                sh["merged"].append(el.get("ref"))
            elif n in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter") and el.text:
                sh["header_footer"].append(f"{n}: {el.text}")
            elif n == "dataValidation":
                f1 = el.find(f"{{{NS_MAIN}}}formula1")
                if f1 is not None and f1.text:
                    sh["validations"].append(f"{el.get('sqref')}: {f1.text}")
        for rtype, target in bk.rels(sh["path"]).values():
            if target not in bk.names:
                continue
            if rtype == "comments":
                for cm in bk.xml(target).iter(f"{{{NS_MAIN}}}comment"):
                    sh["comments"].append({"cell": cm.get("ref"), "text": para_text(cm)})
            elif rtype == "threadedComment":
                for cm in bk.xml(target).iter():
                    if ln(cm.tag) == "threadedComment":
                        sh["comments"].append({"cell": cm.get("ref"), "text": "".join(
                            x.text or "" for x in cm.iter() if ln(x.tag) == "text")})
            elif rtype == "drawing":
                parse_drawing(bk, target, sh, model, img_dir)
    for p in ("docProps/core.xml", "docProps/custom.xml"):
        if p in bk.names:
            for el in bk.xml(p).iter():
                if el.text and el.text.strip() and len(el) == 0:
                    key = ln(el.tag)
                    if p.endswith("custom.xml"):
                        par = next((q for q in bk.xml(p) if el in list(q.iter())), None)
                        key = par.get("name") if par is not None else key
                    if key not in ("lastModifiedBy",):
                        model["properties"][key] = el.text.strip()
    return model


def anchor_cell(anc):
    fr = next((x for x in anc if ln(x.tag) == "from"), None)
    if fr is None:
        return ""
    vals = {ln(x.tag): x.text for x in fr}
    try:
        return f"{col_letter(int(vals['col']))}{int(vals['row'])+1}"
    except Exception:
        return ""


def parse_drawing(bk, dpath, sh, model, img_dir):
    root = bk.xml(dpath)
    drels = bk.rels(dpath)

    def walk(el, cell, group):
        for ch in el:
            n = ln(ch.tag)
            if n not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                continue
            nv = next((x for x in ch.iter() if ln(x.tag) == "cNvPr"), None)
            sid = nv.get("id") if nv is not None else ""
            name = nv.get("name", "") if nv is not None else ""
            alt = " ".join(filter(None, [nv.get("descr", ""), nv.get("title", "")])) if nv is not None else ""
            geom = next((x.get("prst") for x in ch.iter() if ln(x.tag) == "prstGeom"), "")
            if n == "grpSp":
                walk(ch, cell, group + [name])
            elif n == "cxnSp":
                st = next((x.get("id") for x in ch.iter() if ln(x.tag) == "stCxn"), None)
                en = next((x.get("id") for x in ch.iter() if ln(x.tag) == "endCxn"), None)
                head = any(ln(x.tag) == "tailEnd" and x.get("type", "none") != "none" for x in ch.iter())
                tail = any(ln(x.tag) == "headEnd" and x.get("type", "none") != "none" for x in ch.iter())
                if st and en:
                    a, b = (en, st) if (tail and not head) else (st, en)
                    sh["edges"].append({"from": a, "to": b, "arrow": head or tail, "name": name})
            elif n == "sp":
                body = next((x for x in ch if ln(x.tag) == "txBody"), None)
                sh["shapes"].append({"id": sid, "name": name, "type": geom or "shape", "text": para_text(body) if body is not None else "",
                                     "alt": alt, "cell": cell, "group": " > ".join(group)})
            elif n == "pic":
                blip = next((x for x in ch.iter() if ln(x.tag) == "blip"), None)
                rid = blip.get(f"{{{NS_R}}}embed") if blip is not None else None
                if rid in drels and drels[rid][1] in bk.names:
                    add_image(bk, drels[rid][1], model, img_dir, sh["name"], cell, name, alt)
                sh["shapes"].append({"id": sid, "name": name, "type": "picture", "text": "", "alt": alt, "cell": cell,
                                     "group": " > ".join(group), "image_no": len(model["images"])})
            elif n == "graphicFrame":
                for x in ch.iter():
                    if ln(x.tag) == "chart":
                        rid = x.get(f"{{{NS_R}}}id")
                        if rid in drels and drels[rid][1] in bk.names:
                            sh["charts"].append(parse_chart(bk.xml(drels[rid][1]), name, cell))
    for anc in root:
        walk(anc if ln(anc.tag) != "absoluteAnchor" else anc, anchor_cell(anc), [])


def parse_chart(root, name, cell):
    out = {"name": name, "cell": cell, "title": "", "series": [], "axes": []}
    for el in root.iter():
        n = ln(el.tag)
        if n == "title":
            t = para_text(el)
            parent_is_axis = False
            if t and not out["title"]:
                out["title"] = t
            elif t:
                out["axes"].append(t)
        elif n == "ser":
            tx = next((x for x in el if ln(x.tag) == "tx"), None)
            nm = ""
            if tx is not None:
                nm = para_text(tx) or "".join(v.text or "" for v in tx.iter() if ln(v.tag) in ("v", "f"))
            refs = [f.text for f in el.iter() if ln(f.tag) == "f" and f.text]
            out["series"].append({"name": nm, "refs": refs})
    return out


def add_image(bk, part, model, img_dir, sheet, cell, name, alt):
    os.makedirs(img_dir, exist_ok=True)
    data = bk.z.read(part)
    no = len(model["images"]) + 1
    fn = f"img{no}{os.path.splitext(part)[1].lower()}"
    with open(os.path.join(img_dir, fn), "wb") as fh:
        fh.write(data)
    model["images"].append({"no": no, "file": f"images/{model['doc_id']}/{fn}", "sheet": sheet, "cell": cell,
                            "name": name, "alt": alt, "sha1": hashlib.sha1(data).hexdigest(), "text": None})


# ------------------------------------------------------------------ rendering
def grid_bounds(cells):
    rs = [int(k.split(",")[0]) for k in cells]; cs = [int(k.split(",")[1]) for k in cells]
    return (min(rs), max(rs), min(cs), max(cs)) if rs else None


def md_escape(s):
    return str(s).replace("|", "\\|").replace("\n", "<br>")


def render(model, src_rel, kb):
    L = [f"# {model['doc_id']}", "", f"- 元ファイル: `{src_rel}`"]
    if model.get("note"):
        L.append(f"- 注記: {model['note']}")
    for k, v in model["properties"].items():
        L.append(f"- プロパティ {k}: {v}")
    L.append("- シート: " + "、".join(f"{s['name']}" + ("" if s["state"] == "visible" else f"（{s['state']}）") for s in model["sheets"]))
    if model["defined_names"]:
        L.append("- 名前定義: " + "、".join(f"{d['name']}={d['value']}" for d in model["defined_names"]))
    for sh in model["sheets"]:
        L += ["", f"## シート「{sh['name']}」" + ("" if sh["state"] == "visible" else f"（{sh['state']}：非表示シート）")]
        b = grid_bounds(sh["cells"])
        if b:
            r0, r1, c0, c1 = b
            csv_rel = None
            if c1 < 300 and r1 < 200000:
                csv_rel = f"tables/{model['doc_id']}/{safe(sh['name'])}.csv"
                p = os.path.join(kb, csv_rel); os.makedirs(os.path.dirname(p), exist_ok=True)
                with open(p, "w", encoding="utf-8-sig", newline="") as fh:
                    w = csv.writer(fh)
                    for r in range(0, r1 + 1):
                        w.writerow([sh["cells"].get(f"{r},{c}", "") for c in range(0, c1 + 1)])
            n = len(sh["cells"])
            dense = (c1 - c0 + 1) <= 15 and (r1 - r0 + 1) <= 400 and n >= 0.25 * (c1 - c0 + 1) * (r1 - r0 + 1)
            L.append(f"範囲 {col_letter(c0)}{r0+1}:{col_letter(c1)}{r1+1}、値のあるセル {n} 個" +
                     (f"。全データCSV: `{csv_rel}`（1行目=Excelの1行目、A列から）" if csv_rel else ""))
            if sh["hidden_rows"]:
                L.append(f"非表示行: {sh['hidden_rows'][:50]}")
            if sh["hidden_cols"]:
                L.append(f"非表示列: {sh['hidden_cols']}")
            if sh["merged"]:
                L.append(f"結合セル: {', '.join(sh['merged'][:50])}")
            if dense:
                L += ["", "| 行 | " + " | ".join(col_letter(c) for c in range(c0, c1 + 1)) + " |",
                      "|---|" + "---|" * (c1 - c0 + 1)]
                for r in range(r0, r1 + 1):
                    vals = [sh["cells"].get(f"{r},{c}", "") for c in range(c0, c1 + 1)]
                    if any(vals):
                        L.append(f"| {r+1} | " + " | ".join(md_escape(v) for v in vals) + " |")
            else:
                L += ["", "セル（行ごと、値のあるもののみ）:"]
                rows = {}
                for k, v in sh["cells"].items():
                    r, c = map(int, k.split(","))
                    rows.setdefault(r, []).append((c, v))
                for i, r in enumerate(sorted(rows)):
                    if i >= 400:
                        L.append(f"- …（以下省略。全体はCSV参照）"); break
                    L.append(f"- {r+1}行: " + " / ".join(f"{col_letter(c)}{r+1}={md_escape(v)}" for c, v in sorted(rows[r])))
        for title, key in (("ヘッダー/フッター", "header_footer"), ("入力規則", "validations")):
            if sh[key]:
                L += ["", f"### {title}"] + [f"- {x}" for x in sh[key]]
        if sh["comments"]:
            L += ["", "### コメント・メモ"] + [f"- {c['cell']}: {md_escape(c['text'])}" for c in sh["comments"]]
        shapes = [s for s in sh["shapes"] if s["type"] != "picture"]
        if shapes:
            L += ["", "### 図形・テキストボックス"]
            for s in shapes:
                where = f"{s['cell']}付近" + (f"、グループ {s['group']}" if s["group"] else "")
                L.append(f"- 「{s['name']}」（{s['type']}、{where}）: {md_escape(s['text']) or '（文字なし）'}" +
                         (f"［代替テキスト: {s['alt']}］" if s["alt"] else ""))
        if sh["edges"]:
            label = {s["id"]: (s["text"] or s["name"]).replace("\n", " ") for s in sh["shapes"]}
            L += ["", "### フロー（図形をつなぐ矢印・線）"]
            for e in sh["edges"]:
                L.append(f"- {label.get(e['from'], '?')} {'→' if e['arrow'] else '—'} {label.get(e['to'], '?')}")
            nxt = {}
            for e in sh["edges"]:
                nxt.setdefault(label.get(e["from"], "?"), []).append(label.get(e["to"], "?"))
            L.append("- 次の工程一覧: " + "；".join(f"{k} → {'、'.join(v)}" for k, v in nxt.items()))
        for ch in sh["charts"]:
            L += ["", f"### グラフ「{ch['name']}」（{ch['cell']}付近）", f"- タイトル: {ch['title'] or '（なし）'}"]
            if ch["axes"]:
                L.append(f"- 軸タイトル等: {'、'.join(ch['axes'])}")
            for s in ch["series"]:
                L.append(f"- 系列: {s['name']}（参照 {', '.join(s['refs'])}）")
        imgs = [i for i in model["images"] if i["sheet"] == sh["name"]]
        for im in imgs:
            L += ["", f"### 画像{im['no']}（{im['cell']}付近、「{im['name']}」）",
                  f"- ファイル: `{im['file']}`" + (f"、代替テキスト: {im['alt']}" if im["alt"] else ""),
                  f"- 画像内の文字・内容: {im['text'] if im['text'] else PENDING}"]
    out = os.path.join(kb, "docs", f"{model['doc_id']}.md")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")


def write_index(kb, manifest):
    L = ["# ナレッジ索引", "", f"更新: {dt.datetime.now().strftime('%Y/%m/%d %H:%M')}", "",
         "| doc_id | 元ファイル | シート | 画像(未記入) | 更新日時(元) |", "|---|---|---|---|---|"]
    pend = []
    for src, m in sorted(manifest["files"].items()):
        if m.get("status") != "ok":
            continue
        model = json.load(open(os.path.join(kb, "model", f"{m['doc_id']}.json"), encoding="utf-8"))
        sheets = "、".join(s["name"] + ("(非表示)" if s["state"] != "visible" else "") for s in model["sheets"])
        np = [i for i in model["images"] if not i["text"]]
        pend += [(m["doc_id"], i) for i in np]
        L.append(f"| [{m['doc_id']}](docs/{m['doc_id']}.md) | {src} | {sheets} | {len(model['images'])}({len(np)}) | {m['mtime']} |")
    bad = [(s, m) for s, m in manifest["files"].items() if m.get("status") != "ok"]
    if bad:
        L += ["", "## 読めなかったファイル"] + [f"- {s}: {m.get('reason')}" for s, m in bad]
    L += ["", f"## 未記入の画像: {len(pend)} 枚"] + [f"- {d} 画像{i['no']}: `{i['file']}`" for d, i in pend]
    with open(os.path.join(kb, "index.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    return pend


def is_encrypted(head):
    return head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" and "EncryptedPackage".encode("utf-16-le") in head


def build(src, kb, exclude):
    os.makedirs(kb, exist_ok=True)
    mpath = os.path.join(kb, "manifest.json")
    manifest = json.load(open(mpath, encoding="utf-8")) if os.path.exists(mpath) else {"files": {}}
    seen, changed = set(), 0
    kb_abs = os.path.abspath(kb)
    for dp, dns, fns in os.walk(src):
        dns[:] = [d for d in dns if d not in exclude and not d.startswith(".") and os.path.abspath(os.path.join(dp, d)) != kb_abs]
        for fn in sorted(fns):
            if fn.startswith("~$") or not fn.lower().endswith(EXCEL_EXT):
                continue
            path = os.path.join(dp, fn); rel = os.path.relpath(path, src); seen.add(rel)
            stt = os.stat(path)
            sig = f"{stt.st_size}-{int(stt.st_mtime)}"
            old = manifest["files"].get(rel)
            if old and old.get("sig") == sig and (old.get("status") != "ok" or os.path.exists(os.path.join(kb, "model", f"{old['doc_id']}.json"))):
                continue
            doc_id = safe(os.path.splitext(rel)[0].replace(os.sep, "__"))
            entry = {"doc_id": doc_id, "sig": sig, "mtime": dt.datetime.fromtimestamp(stt.st_mtime).strftime("%Y/%m/%d %H:%M")}
            prev = {}
            mp = os.path.join(kb, "model", f"{doc_id}.json")
            if os.path.exists(mp):
                prev = {i["sha1"]: i["text"] for i in json.load(open(mp, encoding="utf-8"))["images"] if i.get("text")}
            try:
                head = open(path, "rb").read()
                if is_encrypted(head):
                    entry.update(status="locked", reason="パスワードで暗号化されていて読めない")
                elif head[:2] == b"PK":
                    model = parse_xlsx(path, doc_id, kb); entry["status"] = "ok"
                else:
                    exe = shutil.which("soffice") or shutil.which("libreoffice")
                    if not exe:
                        entry.update(status="error", reason=".xls だが LibreOffice が無く変換できない")
                    else:
                        with tempfile.TemporaryDirectory() as tmp:
                            s2 = os.path.join(tmp, "in" + os.path.splitext(fn)[1]); shutil.copy(path, s2)
                            subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", tmp, s2],
                                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                            model = parse_xlsx(os.path.join(tmp, "in.xlsx"), doc_id, kb, note=".xls を LibreOffice で一時変換して読み取り")
                        entry["status"] = "ok"
                if entry.get("status") == "ok":
                    for im in model["images"]:
                        im["text"] = prev.get(im["sha1"])
                    os.makedirs(os.path.join(kb, "model"), exist_ok=True)
                    json.dump(model, open(mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
                    shutil.rmtree(os.path.join(kb, "tables", doc_id), ignore_errors=True)
                    render(model, rel, kb)
            except Exception as e:
                entry.update(status="error", reason=f"{type(e).__name__}: {e}")
            manifest["files"][rel] = entry; changed += 1
    for rel in list(manifest["files"]):
        if rel not in seen:
            d = manifest["files"].pop(rel)["doc_id"]
            for sub in ("docs", "model"):
                f = os.path.join(kb, sub, f"{d}.{'md' if sub == 'docs' else 'json'}")
                if os.path.exists(f): os.remove(f)
            for sub in ("tables", "images"):
                shutil.rmtree(os.path.join(kb, sub, d), ignore_errors=True)
            changed += 1
    json.dump(manifest, open(mpath, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    pend = write_index(kb, manifest)
    ok = sum(1 for m in manifest["files"].values() if m.get("status") == "ok")
    print(f"files={len(manifest['files'])} ok={ok} updated={changed} pending_images={len(pend)}")
    print(f"index: {os.path.join(kb, 'index.md')}")


def set_image(kb, doc_id, no, text):
    mp = os.path.join(kb, "model", f"{doc_id}.json")
    model = json.load(open(mp, encoding="utf-8"))
    im = next(i for i in model["images"] if i["no"] == int(no))
    im["text"] = "（文字なし）" if text.strip() == "-" else text.strip()
    json.dump(model, open(mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    manifest = json.load(open(os.path.join(kb, "manifest.json"), encoding="utf-8"))
    rel = next(r for r, m in manifest["files"].items() if m["doc_id"] == doc_id)
    render(model, rel, kb)
    pend = write_index(kb, manifest)
    print(f"ok: {doc_id} 画像{no}  残り未記入 {len(pend)} 枚")


def main():
    a = sys.argv[1:]
    if not a:
        print(__doc__); sys.exit(1)
    if a[0] == "build":
        ex = a[a.index("--exclude") + 1:] if "--exclude" in a else []
        build(a[1], a[2], ex)
    elif a[0] == "pending":
        manifest = json.load(open(os.path.join(a[1], "manifest.json"), encoding="utf-8"))
        for d, i in write_index(a[1], manifest):
            print(f"{d}\t{i['no']}\t{os.path.join(a[1], i['file'])}\t{i['sheet']}\t{i['cell']}\t{i['alt']}")
    elif a[0] == "set-image":
        set_image(a[1], a[2], a[3], " ".join(a[4:]))
    else:
        print(__doc__); sys.exit(1)


if __name__ == "__main__":
    main()
