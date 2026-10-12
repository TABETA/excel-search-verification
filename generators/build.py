"""Generate ~100 Excel test files for 'can Cowork search inside Excel' experiment.

Each file gets exactly one scenario. answer_key.csv records where the keyword was planted.
Keyword: ペリカン承認
"""
import csv, io, json, os, random, re, shutil, sys, zipfile
from xml.sax.saxutils import escape
import xlsxwriter
import images as IM

KW = "ペリカン承認"
ROOT = os.path.abspath(os.environ.get("WORK_DIR", "."))
OUT = f"{ROOT}/out/xlsx_stage"
IMG = f"{ROOT}/out/img"
rng = random.Random(20261012)

# --------------------------------------------------------------------------- filler
ITEMS = ["ブラケット", "シャフト", "カバー", "ギア", "ベアリング", "スペーサー", "ハウジング", "ボルトM6",
         "ナットM8", "Oリング", "センサ台座", "ケーブル", "ヒンジ", "プーリー", "フランジ", "ガスケット"]
PEOPLE = ["田中", "鈴木", "佐藤", "高橋", "伊藤", "渡辺", "山本", "中村", "小林", "加藤"]
STATUS = ["完了", "仕掛", "保留", "確認済", "要再検査", "出荷可", "差戻し"]
TEMPLATES = [
    ("品番", "品名", "数量", "単価", "担当"),
    ("No", "点検項目", "基準値", "測定値", "判定"),
    ("日付", "工程", "作業者", "出来高", "備考"),
    ("案件No", "顧客", "品名", "納期", "状態"),
]
DECOY_SHAPE_TEXT = ["確認済", "要確認", "担当：田中", "暫定対策", "部長決裁", "注意：治具交換",
                    "改訂3で変更", "出荷前に再検査"]


def filler(ws, rows=None, start_row=0, wb=None):
    tpl = rng.choice(TEMPLATES)
    rows = rows or rng.randint(12, 40)
    for c, h in enumerate(tpl):
        ws.write(start_row, c, h)
    for r in range(1, rows + 1):
        for c, h in enumerate(tpl):
            if h in ("数量", "出来高"):
                v = rng.randint(1, 500)
            elif h in ("単価", "測定値", "基準値"):
                v = round(rng.uniform(0.5, 999), 2)
            elif h in ("品名", "点検項目"):
                v = rng.choice(ITEMS)
            elif h in ("担当", "作業者"):
                v = rng.choice(PEOPLE)
            elif h in ("判定", "状態", "備考"):
                v = rng.choice(STATUS)
            elif h in ("日付", "納期"):
                v = f"2026/{rng.randint(1,12):02d}/{rng.randint(1,28):02d}"
            elif h == "顧客":
                v = rng.choice(["A社", "B工業", "C製作所", "D電機"])
            elif h == "工程":
                v = rng.choice(["切削", "研磨", "組立", "検査", "梱包"])
            else:
                v = f"{h[:1]}{start_row + r:04d}" if h != "No" else r
            ws.write(start_row + r, c, v)
    ws.set_column(0, len(tpl) - 1, 13)
    return rows


def decorate(wb, ws, tag):
    """Random non-keyword shapes/images/comments so every file looks 'real'."""
    if rng.random() < 0.45:
        ws.insert_textbox(f"H{rng.randint(2,10)}", rng.choice(DECOY_SHAPE_TEXT),
                          {"width": 160, "height": 50})
    if rng.random() < 0.35:
        p = f"{IMG}/{tag}_decoy.png"
        IM.decoy(p, rng)
        ws.insert_image(f"H{rng.randint(12,20)}", p, {"description": "図", "x_scale": 0.5, "y_scale": 0.5})
    if rng.random() < 0.2:
        ws.write_comment("B3", rng.choice(["測定器No.12使用", "前回値と比較のこと", "要フォロー"]))


# --------------------------------------------------------------------------- custom DrawingML
EMU = 9525


def _run(text, sz=1100, color="000000", b=False, lang="ja-JP"):
    return (f'<a:r><a:rPr lang="{lang}" altLang="en-US" sz="{sz}"{" b=\"1\"" if b else ""}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr><a:t>{escape(text)}</a:t></a:r>')


def _para(runs, algn="ctr"):
    if isinstance(runs, str):
        runs = [_run(runs)]
    return f'<a:p><a:pPr algn="{algn}"/>{"".join(runs)}</a:p>'


def sp(sid, name, paras, x, y, cx, cy, prst="rect", vert=None, fill="DDEBF7"):
    if isinstance(paras, str):
        paras = [_para(paras)]
    v = f' vert="{vert}"' if vert else ""
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{sid}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="2F5597"/></a:solidFill></a:ln></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"{v}/><a:lstStyle/>'
            f'{"".join(paras)}</xdr:txBody></xdr:sp>')


def grp(gid, name, children, x, y, cx, cy):
    return (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{gid}" name="{name}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
            f'<xdr:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
            f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></xdr:grpSpPr>'
            f'{"".join(children)}</xdr:grpSp>')


def anchor(col, row, col2, row2, inner):
    def pt(c, r):
        return f"<xdr:col>{c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>0</xdr:rowOff>"
    return (f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from>{pt(col,row)}</xdr:from><xdr:to>{pt(col2,row2)}</xdr:to>'
            f'{inner}<xdr:clientData/></xdr:twoCellAnchor>')


def pos(col, row):
    return col * 64 * EMU, row * 20 * EMU


PENDING = {}  # path -> {placeholder: anchor_xml}


def placeholder(ws, path, cell_col, cell_row, anchor_xml):
    ph = f"__PH{len(PENDING.get(path, {}))}_{rng.randint(1000,9999)}__"
    ws.insert_textbox(cell_row, cell_col, ph)
    PENDING.setdefault(path, {})[ph] = anchor_xml


def postprocess(path):
    reps = PENDING.get(path)
    if not reps:
        return
    zin = zipfile.ZipFile(path)
    buf = io.BytesIO()
    zout = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    done = set()
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("xl/drawings/drawing") and item.filename.endswith(".xml"):
            s = data.decode("utf-8")

            def sub(m):
                for ph, xml in reps.items():
                    if ph in m.group(0):
                        done.add(ph)
                        return xml
                return m.group(0)
            s = re.sub(r"<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>", sub, s, flags=re.S)
            data = s.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    assert done == set(reps), (path, set(reps) - done)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue())


# --------------------------------------------------------------------------- scenarios
SCEN = []  # (category, expected, builder) ; builder(wb, path, tag) -> location string ; or special


def scen(category, expected, n=1):
    def deco(fn):
        for i in range(n):
            SCEN.append((category, expected, fn, i))
        return fn
    return deco


def new_wb(path):
    return xlsxwriter.Workbook(path)


# ---- cells -----------------------------------------------------------------
@scen("セル_通常", "HIT", 4)
def cell_plain(wb, path, tag, i):
    ws = wb.add_worksheet("一覧")
    n = filler(ws)
    if i == 0:
        r = rng.randint(3, n); ws.write(r, 4, KW); return f"一覧!E{r+1}（セル値そのもの）"
    if i == 1:
        r = n + 3; ws.write(r, 0, f"申し送り：本件は{KW}の取得後に量産移行とする。"); return f"一覧!A{r+1}（文中）"
    if i == 2:
        fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        ws.merge_range(n + 3, 0, n + 5, 4, f"{KW}フロー適用", fmt); return f"一覧!A{n+4}:E{n+6}（結合セル）"
    ws2 = wb.add_worksheet("集計")
    m = filler(ws2, rows=60)
    ws2.write(m + 2, 2, f"最終区分={KW}"); return f"集計!C{m+3}（2枚目シートの下方）"


@scen("セル_書式混在(リッチテキスト)", "HIT", 3)
def cell_rich(wb, path, tag, i):
    ws = wb.add_worksheet("記録")
    n = filler(ws)
    b = wb.add_format({"bold": True}); red = wb.add_format({"font_color": "red"}); u = wb.add_format({"underline": 1})
    parts = [[b, "ペリ", red, "カン承認"], [red, "ペリカン", b, "承認", "（済）"], [u, "ペ", b, "リカン承", red, "認"]][i]
    if i == 1:
        ws.write_rich_string(n + 2, 1, red, "ペリカン", b, "承認", "（済）")
    else:
        ws.write_rich_string(n + 2, 1, *parts)
    return f"記録!B{n+3}（1語が書式違いの複数ランに分割）"


@scen("セル_数式結果", "HIT", 2)
def cell_formula(wb, path, tag, i):
    ws = wb.add_worksheet("計算")
    n = filler(ws)
    if i == 0:
        ws.write_formula(n + 2, 1, '="ペリカン"&"承認"', None, KW)
        return f'計算!B{n+3}（数式 ="ペリカン"&"承認" のキャッシュ値）'
    ws.write(n + 2, 0, "ペリカン"); ws.write(n + 2, 1, "承認")
    ws.write_formula(n + 2, 3, f"=A{n+3}&B{n+3}", None, KW)
    return f"計算!D{n+3}（=A&B の結果。A,B単体は部分語）"


@scen("セル_非表示シート", "HIT", 3)
def cell_hidden_sheet(wb, path, tag, i):
    ws = wb.add_worksheet("表紙"); filler(ws)
    h = wb.add_worksheet(["マスタ", "旧データ", "設定"][i]); filler(h, rows=10)
    h.write(14, 1, f"区分コード：{KW}")
    if i == 2:
        h.very_hidden(); return "設定!B15（veryHidden シート）"
    h.hide(); return f"{h.name}!B15（非表示シート）"


@scen("セル_非表示行列", "HIT", 2)
def cell_hidden_rowcol(wb, path, tag, i):
    ws = wb.add_worksheet("台帳"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 0, KW); ws.set_row(n + 2, None, None, {"hidden": True}); return f"台帳!A{n+3}（非表示行）"
    ws.write(5, 10, KW); ws.set_column(10, 10, None, None, {"hidden": True}); return "台帳!K6（非表示列）"


@scen("セル_白文字", "HIT", 2)
def cell_white(wb, path, tag, i):
    ws = wb.add_worksheet("帳票"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 3, KW, wb.add_format({"font_color": "white"})); return f"帳票!D{n+3}（白文字）"
    ws.write(n + 2, 3, KW, wb.add_format({"font_color": "#FFF2CC", "bg_color": "#FFF2CC"})); return f"帳票!D{n+3}（背景色と同色の文字）"


@scen("セル_遠方セル", "HIT", 2)
def cell_far(wb, path, tag, i):
    ws = wb.add_worksheet("データ"); filler(ws)
    if i == 0:
        ws.write(1048575, 16383, KW); return "データ!XFD1048576（最終セル）"
    ws.write(20000, 50, KW); return "データ!AY20001"


@scen("コメント(メモ)", "HIT", 4)
def comment(wb, path, tag, i):
    ws = wb.add_worksheet("チェック"); filler(ws)
    txt = [KW, f"この行は{KW}が必要", f"確認者：田中\n経緯：設計変更に伴い\n{KW}を追加取得", f"{KW}"][i]
    opt = {"visible": True} if i == 3 else {}
    ws.write_comment("C5", txt, opt)
    return "チェック!C5 のメモ" + ("（常時表示）" if i == 3 else "")


@scen("ヘッダー/フッター", "HIT", 2)
def header_footer(wb, path, tag, i):
    ws = wb.add_worksheet("印刷用"); filler(ws)
    if i == 0:
        ws.set_header(f"&C{KW}要"); return "印刷用 シートのヘッダー（中央）"
    ws.set_footer(f"&R{KW} 改訂2 &P/&N"); return "印刷用 シートのフッター（右）"


@scen("入力規則/名前定義", "HIT", 2)
def validation_name(wb, path, tag, i):
    ws = wb.add_worksheet("入力"); n = filler(ws)
    if i == 0:
        ws.data_validation(1, 5, n, 5, {"validate": "list", "source": ["未着手", KW, "却下"]})
        ws.write(0, 5, "区分")
        return "入力!F2:F の入力規則（リスト選択肢）"
    wb.define_name("承認ルート", f'="{KW}"'); return "名前定義『承認ルート』の値"


@scen("ブックのプロパティ", "HIT", 2)
def docprops(wb, path, tag, i):
    ws = wb.add_worksheet("本文"); filler(ws)
    if i == 0:
        wb.set_properties({"title": f"{KW}手順", "author": "品証"}); return "ファイルのプロパティ：タイトル"
    wb.set_custom_property("承認区分", KW); return "ファイルのプロパティ：ユーザー設定プロパティ"


@scen("シート名", "HIT", 2)
def sheetname(wb, path, tag, i):
    a = wb.add_worksheet("概要"); filler(a)
    b = wb.add_worksheet([KW, f"{KW}(2)"][i]); filler(b, rows=8)
    return f"シート名『{b.name}』"


# ---- shapes ----------------------------------------------------------------
@scen("図形_テキストボックス", "HIT", 5)
def textbox(wb, path, tag, i):
    ws = wb.add_worksheet("説明"); filler(ws)
    if i == 4:
        ws2 = wb.add_worksheet("補足"); filler(ws2, rows=5)
        ws2.insert_textbox("B10", f"注意：{KW}が未取得の場合は出荷不可", {"width": 300, "height": 60})
        return "補足 シートのテキストボックス"
    txt = [KW, f"本工程は{KW}対象", f"変更点\n1. 治具変更\n2. {KW}追加\n3. 記録様式改訂",
           f"{KW}（小さい字）"][i]
    opt = {"width": 260, "height": 90 if i == 2 else 50}
    if i == 3:
        opt["font"] = {"size": 6}
    ws.insert_textbox("H3", txt, opt)
    return "説明 シートのテキストボックス"


@scen("図形_オートシェイプ", "HIT", 5)
def autoshape(wb, path, tag, i):
    ws = wb.add_worksheet("フロー"); filler(ws)
    prst, vert, label = [("flowChartProcess", None, "フローチャート:処理"),
                         ("wedgeRectCallout", None, "吹き出し"),
                         ("rightArrow", None, "右矢印"),
                         ("ellipse", None, "楕円"),
                         ("flowChartDecision", "eaVert", "ひし形・縦書き")][i]
    x, y = pos(7, 3)
    xml = anchor(7, 3, 11, 9, sp(101, f"図形 {i+2}", KW if i != 1 else f"ここで{KW}", x, y, 4 * 64 * EMU, 6 * 20 * EMU, prst, vert))
    placeholder(ws, path, 7, 3, xml)
    return f"フロー シートの図形（{label}）"


@scen("図形_グループ化", "HIT", 4)
def group(wb, path, tag, i):
    ws = wb.add_worksheet("構成"); filler(ws)
    x, y = pos(7, 2); w, h = 120 * EMU * 2, 40 * EMU
    a = sp(201, "四角形 1", "受入", x, y, w, h)
    b = sp(202, "四角形 2", "加工", x, y + h * 2, w, h)
    c = sp(203, "四角形 3", KW, x, y + h * 4, w, h, "roundRect")
    d = sp(204, "四角形 4", "出荷", x, y + h * 6, w, h)
    if i == 0:
        inner = grp(200, "グループ化 1", [a, c], x, y, w, h * 5)
    elif i == 1:
        inner = grp(200, "グループ化 1", [a, grp(210, "グループ化 2", [b, c], x, y + h * 2, w, h * 3)], x, y, w, h * 5)
    elif i == 2:
        inner = grp(200, "グループ化 1", [a, grp(210, "グループ化 2", [b, grp(220, "グループ化 3", [c, d], x, y + h * 4, w, h * 3)], x, y + h * 2, w, h * 5)], x, y, w, h * 7)
    else:
        many = [sp(300 + k, f"四角形 {k}", f"工程{k}", x, y + k * h, w, h) for k in range(8)]
        many.append(sp(399, "四角形 99", KW, x, y + 8 * h, w, h))
        inner = grp(200, "グループ化 1", many, x, y, w, h * 9)
    depth = ["1階層", "2階層", "3階層", "9個中の最後"][i]
    placeholder(ws, path, 7, 2, anchor(7, 2, 11, 16, inner))
    return f"構成 シートのグループ図形内（{depth}）"


@scen("図形_ラン分割", "HIT", 3)
def shape_runs(wb, path, tag, i):
    ws = wb.add_worksheet("注記"); filler(ws)
    runs = [[_run("ペリ", b=True), _run("カン承認")],
            [_run("ペリカン", color="C00000"), _run("承認", sz=1400, b=True)],
            [_run("ペ", lang="en-US"), _run("リカン承", lang="ja-JP"), _run("認", lang="ja-JP", b=True)]][i]
    x, y = pos(7, 4)
    placeholder(ws, path, 7, 4, anchor(7, 4, 11, 8, sp(101, "テキスト 1", [_para(runs)], x, y, 4 * 64 * EMU, 4 * 20 * EMU)))
    return "注記 シートの図形（1語が書式違いの複数ランに分割）"


@scen("図形_見えにくい位置", "HIT", 2)
def shape_hidden(wb, path, tag, i):
    ws = wb.add_worksheet("本体"); filler(ws)
    if i == 0:
        x, y = pos(60, 300)
        placeholder(ws, path, 60, 300, anchor(60, 300, 64, 304, sp(101, "四角形 1", KW, x, y, 4 * 64 * EMU, 4 * 20 * EMU)))
        return "本体 シートの BI301 付近の図形（画面外）"
    h = wb.add_worksheet("作業用"); filler(h, rows=5); h.hide()
    x, y = pos(2, 8)
    placeholder(h, path, 2, 8, anchor(2, 8, 6, 12, sp(101, "四角形 1", KW, x, y, 4 * 64 * EMU, 4 * 20 * EMU)))
    return "非表示シート『作業用』の図形"


# ---- charts ----------------------------------------------------------------
@scen("グラフ", "HIT", 4)
def chart(wb, path, tag, i):
    ws = wb.add_worksheet("推移")
    ws.write_row(0, 0, ["月", "件数", "不良"])
    for r in range(1, 13):
        ws.write_row(r, 0, [f"{r}月", rng.randint(20, 90), rng.randint(0, 9)])
    ch = wb.add_chart({"type": "column" if i % 2 == 0 else "line"})
    sname = f"{KW}件数" if i == 2 else "件数"
    ch.add_series({"name": sname, "categories": "=推移!$A$2:$A$13", "values": "=推移!$B$2:$B$13"})
    if i == 0:
        ch.set_title({"name": f"{KW}の月別推移"}); loc = "グラフタイトル"
    elif i == 1:
        ch.set_title({"name": "月別推移"}); ch.set_y_axis({"name": f"{KW}数"}); loc = "グラフの縦軸ラベル"
    elif i == 2:
        ch.set_title({"name": "月別推移"}); loc = "グラフの系列名（文字列直書き）"
    else:
        ch.set_title({"name": f"{KW}状況"}); cs = wb.add_chartsheet("グラフ1"); cs.set_chart(ch)
        return "グラフシート『グラフ1』のタイトル"
    ws.insert_chart("E2", ch)
    return f"推移 シートの{loc}"


# ---- images ----------------------------------------------------------------
PIX = list(IM.PIXEL_STYLES)


@scen("画像_画素内の文字", "HIT", 10)
def image_pixels(wb, path, tag, i):
    ws = wb.add_worksheet("資料"); filler(ws)
    style = PIX[i]
    p = f"{IMG}/{tag}.{'jpg' if style == 'noisy' else 'png'}"
    IM.PIXEL_STYLES[style](KW, p)
    ws.insert_image("H3", p, {"description": "図1"})
    return f"資料 シートの画像（画素内の文字・スタイル:{style}）"


@scen("画像_代替テキスト", "HIT", 3)
def image_alt(wb, path, tag, i):
    ws = wb.add_worksheet("写真"); filler(ws)
    p = f"{IMG}/{tag}.png"; IM.photo_like(p, rng)
    ws.insert_image("H3", p, {"description": [KW, f"{KW}済みの現品写真", f"現品写真（{KW}待ち）"][i]})
    return "写真 シートの画像の代替テキスト（画像自体に文字なし）"


# ---- multi -----------------------------------------------------------------
@scen("複数箇所", "HIT", 3)
def multi(wb, path, tag, i):
    ws = wb.add_worksheet("まとめ"); n = filler(ws)
    locs = []
    if i in (0, 2):
        ws.write(n + 2, 1, KW); locs.append(f"まとめ!B{n+3}")
    if i in (0, 1):
        ws.insert_textbox("H2", f"{KW}要", {"width": 200, "height": 40}); locs.append("まとめ シートのテキストボックス")
    if i in (1, 2):
        p = f"{IMG}/{tag}.png"; IM.diagram(KW, p); ws.insert_image("H10", p, {"description": "図1"})
        locs.append("まとめ シートの画像（画素内）")
    if i == 2:
        ws.write_comment("A1", f"{KW}について要確認"); locs.append("まとめ!A1 のメモ")
    return " / ".join(locs)


# ---- variants (表記ゆれ) ----------------------------------------------------
@scen("表記ゆれ", "VARIANT", 5)
def variant(wb, path, tag, i):
    ws = wb.add_worksheet("一覧"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 1, "ﾍﾟﾘｶﾝ承認"); return f"一覧!B{n+3}『ﾍﾟﾘｶﾝ承認』（半角カナ）"
    if i == 1:
        ws.write(n + 2, 1, "ペリカン\n承認", wb.add_format({"text_wrap": True})); return f"一覧!B{n+3}（セル内改行で分断）"
    if i == 2:
        x, y = pos(7, 3)
        placeholder(ws, path, 7, 3, anchor(7, 3, 11, 8, sp(101, "四角形 1", [_para("ペリカン"), _para("承認")], x, y, 4 * 64 * EMU, 5 * 20 * EMU)))
        return "一覧 シートの図形（段落で『ペリカン』『承認』に分断）"
    if i == 3:
        ws.write(n + 2, 1, "ペリカン　承認"); return f"一覧!B{n+3}（間に全角スペース）"
    ws.write(n + 2, 1, "ぺりかん承認"); return f"一覧!B{n+3}（ひらがな）"


# ---- traps (should NOT match) ----------------------------------------------
@scen("紛らわしい語(非該当)", "NO", 5)
def trap(wb, path, tag, i):
    ws = wb.add_worksheet("一覧"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 1, "ペリカン承諾"); return f"一覧!B{n+3}『ペリカン承諾』"
    if i == 1:
        ws.insert_textbox("H3", "ペリカンの承認", {"width": 200, "height": 40}); return "テキストボックス『ペリカンの承認』"
    if i == 2:
        ws.write(n + 2, 1, "ペリカン"); ws.write(n + 2, 2, "承認"); return f"一覧!B{n+3}『ペリカン』とC{n+3}『承認』が別セル"
    if i == 3:
        p = f"{IMG}/{tag}.png"; IM.clear("ペリカン承諾", p); ws.insert_image("H3", p, {"description": "図1"})
        return "画像（画素内）『ペリカン承諾』"
    ws.write(n + 2, 1, "承認ペリカン"); return f"一覧!B{n+3}『承認ペリカン』（語順逆）"


@scen("無関係(ノイズ)", "NO", 15)
def noise(wb, path, tag, i):
    for k in range(rng.randint(1, 3)):
        ws = wb.add_worksheet(["Sheet1", "集計", "明細", "グラフ用"][k])
        filler(ws)
        decorate(wb, ws, f"{tag}_{k}")
    return "（キーワードなし）"


# ---- formats (built from xlsx, converted later) ----------------------------
@scen("形式_xls(旧形式)", "HIT", 3)
def fmt_xls(wb, path, tag, i):
    ws = wb.add_worksheet("Sheet1"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 1, KW); return f"Sheet1!B{n+3}（.xls）"
    if i == 1:
        ws.insert_textbox("H3", f"{KW}要", {"width": 200, "height": 50}); return "Sheet1 のテキストボックス（.xls）"
    ws.write_comment("C4", f"{KW}"); return "Sheet1!C4 のメモ（.xls）"


@scen("形式_パスワード暗号化", "LOCKED", 2)
def fmt_enc(wb, path, tag, i):
    ws = wb.add_worksheet("Sheet1"); n = filler(ws)
    if i == 0:
        ws.write(n + 2, 1, KW); return f"Sheet1!B{n+3}（読み取りパスワード付き。PW=pelican）"
    ws.insert_textbox("H3", KW, {"width": 200, "height": 50}); return "テキストボックス（読み取りパスワード付き。PW=pelican）"


@scen("形式_シート保護", "HIT", 1)
def fmt_protect(wb, path, tag, i):
    ws = wb.add_worksheet("Sheet1"); n = filler(ws)
    ws.write(n + 2, 1, KW); ws.protect("pelican"); return f"Sheet1!B{n+3}（シート保護あり・ファイルは開ける）"


# --------------------------------------------------------------------------- names
DOC = ["見積書", "発注書", "作業手順書", "設備点検表", "品質記録", "工程表", "不具合報告", "会議資料",
       "仕様一覧", "部品表", "在庫表", "出荷記録", "検査成績書", "改善提案", "教育記録", "日報",
       "設計レビュー", "変更管理表", "予算表", "業務フロー"]


def main():
    for d in (OUT, IMG):
        shutil.rmtree(d, ignore_errors=True); os.makedirs(d)
    order = SCEN[:]
    rng.shuffle(order)
    used = set(); names = []
    for _ in order:
        while True:
            nm = f"{rng.choice(DOC)}_{rng.randint(100, 999)}"
            if nm not in used:
                used.add(nm); names.append(nm); break
    key = []
    for (cat, exp, fn, i), nm in zip(order, names):
        tag = nm
        ext_final = ".xls" if fn is fmt_xls else ".xlsx"
        path = f"{OUT}/{nm}.xlsx"
        wb = new_wb(path)
        loc = fn(wb, path, tag, i)
        wb.close()
        postprocess(path)
        key.append({"file": nm + ext_final, "expected": exp, "category": cat, "location": loc,
                    "_convert": "xls" if fn is fmt_xls else ("enc" if fn is fmt_enc else "")})
    with open(f"{ROOT}/out/key_raw.json", "w") as fh:
        json.dump(key, fh, ensure_ascii=False, indent=1)
    print(len(key), "files")


if __name__ == "__main__":
    main()
