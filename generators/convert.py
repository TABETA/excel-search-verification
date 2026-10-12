"""Convert xls targets and encrypt LOCKED targets via LibreOffice UNO."""
import json, os, subprocess, time, uno
from com.sun.star.beans import PropertyValue

ROOT = os.path.join(os.path.abspath(os.environ.get("WORK_DIR", ".")), "out")
STAGE = f"{ROOT}/xlsx_stage"
FINAL = f"{ROOT}/final"


def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


def main():
    os.makedirs(FINAL, exist_ok=True)
    key = json.load(open(f"{ROOT}/key_raw.json"))
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                             "--accept=socket,host=localhost,port=2002;urp;"])
    ctx = None
    for _ in range(60):
        try:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    for k in key:
        stem = k["file"].rsplit(".", 1)[0]
        src = f"{STAGE}/{stem}.xlsx"
        if not k["_convert"]:
            os.replace(src, f"{FINAL}/{k['file']}")
            continue
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
        if k["_convert"] == "xls":
            doc.storeToURL(uno.systemPathToFileUrl(f"{FINAL}/{k['file']}"), (pv("FilterName", "MS Excel 97"),))
        else:
            doc.storeToURL(uno.systemPathToFileUrl(f"{FINAL}/{k['file']}"),
                           (pv("FilterName", "Calc Office Open XML"), pv("Password", "pelican")))
        doc.close(True)
        print("converted", k["file"], k["_convert"])
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)


if __name__ == "__main__":
    main()
