import os, json, zipfile, re, sys, xml.etree.ElementTree as ET, warnings
warnings.filterwarnings("ignore")
import openpyxl
KW="ペリカン承認"
key=json.load(open(os.path.join(os.environ.get("WORK_DIR", "."), "out", "key_raw.json")))
D=os.path.join(os.environ.get("WORK_DIR", "."), "out", "final")
bad=0
for k in key:
    p=f"{D}/{k['file']}"
    if k["_convert"]:
        continue
    z=zipfile.ZipFile(p); raw=False; joined=False
    for n in z.namelist():
        if not (n.endswith(".xml") or n.endswith(".vml") or n.endswith(".rels")): continue
        s=z.read(n).decode("utf-8")
        ET.fromstring(s)  # well-formed
        if KW in s: raw=True
        t=re.sub(r"</a:p>|</si>|</c>|</xdr:sp>","\n",s)
        t=re.sub(r"<[^>]+>","",t)
        if KW in t.replace("\n","\n"): joined=True
        # paragraph-joined (text across runs inside one paragraph)
        for para in re.findall(r"<a:p>.*?</a:p>|<si>.*?</si>",s,flags=re.S):
            if KW in "".join(re.findall(r"<(?:a:)?t[^>]*>([^<]*)</(?:a:)?t>",para)): joined=True
    try:
        wb=openpyxl.load_workbook(p); ok="ok"
        cells=any(KW in str(c.value) for ws in wb.worksheets for c in ws._cells.values() if c.value is not None)
    except Exception as e:
        ok=f"ERR {e}"; cells=None
    print(f"{k['expected']:7} raw={int(raw)} para={int(joined)} opx_cells={int(bool(cells))} {ok:3} {k['category']:22} {k['file']}")
