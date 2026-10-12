"""Generate a realistic fictional-manufacturer Excel document set + Q&A answer key."""
import datetime as dt, json, os, random, shutil, subprocess
import xlsxwriter
from PIL import Image, ImageDraw
import build as B          # drawing helpers (sp, anchor, grp, placeholder, postprocess ...)
import images as IM

_W = os.path.abspath(os.environ.get("WORK_DIR", "."))
OUT = os.path.join(_W, "out", "qa_docs")
KEY = os.path.join(_W, "out", "qa_key.json")
TMP = os.path.join(_W, "out", "qa_tmp")
rng = random.Random(77)
EMU = B.EMU


def wb_new(name):
    p = f"{OUT}/{name}.xlsx"
    return xlsxwriter.Workbook(p), p


def finish(wb, p):
    wb.close(); B.postprocess(p)


def fmts(wb):
    return dict(h=wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1}),
                c=wb.add_format({"border": 1}),
                d=wb.add_format({"border": 1, "num_format": "yyyy/mm/dd"}),
                n=wb.add_format({"border": 1, "num_format": "#,##0"}),
                t=wb.add_format({"bold": True, "font_size": 14}),
                w=wb.add_format({"border": 1, "text_wrap": True, "valign": "top"}))


CUST = ["A社", "B工業", "C製作所", "D電機", "E産業"]
PROD = {"GX-200": ("ギアユニット", 48000), "HS-110": ("シャフト組品", 12500),
        "PL-50": ("プレート", 3200), "BR-30": ("ブラケット", 1800)}


# ---------------------------------------------------------------- orders
def orders(name, start, months, prefix, cancel_note=True):
    wb, p = wb_new(name); F = fmts(wb)
    ws = wb.add_worksheet("受注")
    ws.write(0, 0, f"受注一覧（{name.split('_')[1]}）", F["t"])
    ws.write(1, 0, "※状態が「キャンセル」の行は売上集計の対象外とする。金額は税抜。")
    hdr = ["受注No", "受注日", "顧客", "品番", "品名", "数量", "単価", "金額", "状態"]
    ws.write_row(3, 0, hdr, F["h"])
    rows = []
    r = 4
    n = 0
    for m in range(months):
        y, mo = start[0] + (start[1] - 1 + m) // 12, (start[1] - 1 + m) % 12 + 1
        for _ in range(rng.randint(18, 26)):
            n += 1
            d = dt.date(y, mo, rng.randint(1, 28))
            c = rng.choice(CUST); pn = rng.choice(list(PROD)); q = rng.choice([10, 20, 30, 50, 100, 200])
            if pn == "GX-200": q = rng.choice([2, 4, 5, 8, 10])
            unit = PROD[pn][1]; st = "キャンセル" if rng.random() < 0.07 else rng.choice(["出荷済", "出荷済", "生産中"])
            rows.append(dict(no=f"{prefix}-{n:04d}", date=d, cust=c, pn=pn, qty=q, unit=unit, amt=q * unit, st=st))
    rows.sort(key=lambda x: x["date"])
    for i, x in enumerate(rows):
        x["no"] = f"{prefix}-{i+1:04d}"
    for x in rows:
        ws.write(r, 0, x["no"], F["c"]); ws.write_datetime(r, 1, dt.datetime.combine(x["date"], dt.time()), F["d"])
        ws.write(r, 2, x["cust"], F["c"]); ws.write(r, 3, x["pn"], F["c"]); ws.write(r, 4, PROD[x["pn"]][0], F["c"])
        ws.write(r, 5, x["qty"], F["n"]); ws.write(r, 6, x["unit"], F["n"])
        ws.write_formula(r, 7, f"=F{r+1}*G{r+1}", F["n"], x["amt"]); ws.write(r, 8, x["st"], F["c"])
        r += 1
    ws.set_column(0, 0, 12); ws.set_column(1, 1, 11); ws.set_column(2, 8, 10)
    ws.autofilter(3, 0, r - 1, 8); ws.freeze_panes(4, 0)
    finish(wb, p)
    return rows


# ---------------------------------------------------------------- approval rules
def approval(name, ver, date, bucho_max, hidden_hist):
    wb, p = wb_new(name); F = fmts(wb)
    ws = wb.add_worksheet("権限表")
    ws.write(0, 0, "職務権限規程 別表：決裁権限一覧", F["t"])
    ws.write(1, 0, f"第{ver}版　{date} 施行")
    ws.write_row(3, 0, ["区分", "課長", "部長", "工場長"], F["h"])
    data = [("設備購入", "50万円以下", f"50万円超〜{bucho_max}万円以下", f"{bucho_max}万円超"),
            ("消耗品購入", "30万円以下", "30万円超〜100万円以下", "100万円超"),
            ("外注加工", "100万円以下", "100万円超〜500万円以下", "500万円超"),
            ("国内出張", "日帰り", "宿泊3泊まで", "宿泊4泊以上")]
    for i, row in enumerate(data):
        ws.write_row(4 + i, 0, row, F["c"])
    ws.write(9, 0, "※金額は1件あたり（税抜）。工場長決裁を超えるものは取締役会付議。")
    ws.set_column(0, 3, 24)
    if hidden_hist:
        h = wb.add_worksheet("改訂履歴")
        h.write_row(0, 0, ["版", "施行日", "改訂内容"], F["h"])
        h.write_row(1, 0, ["第4版", "2024/04/01", "外注加工の区分を新設"], F["c"])
        h.write_row(2, 0, ["第5版", "2026/04/01", "設備購入の部長決裁上限を200万円から300万円に引き上げ"], F["c"])
        h.set_column(2, 2, 60); h.hide()
    finish(wb, p)


# ---------------------------------------------------------------- BOM / stock
def bom():
    wb, p = wb_new("部品表_GX-200"); F = fmts(wb)
    ws = wb.add_worksheet("BOM")
    ws.write(0, 0, "部品表　GX-200 ギアユニット（Rev.C）", F["t"])
    ws.write_row(2, 0, ["部品番号", "部品名", "員数", "材質", "仕入先", "リードタイム(日)"], F["h"])
    parts = [("G-101", "入力ギア", 1, "SCM420", "社内加工", 7), ("G-102", "出力ギア", 1, "SCM420", "社内加工", 7),
             ("S-201", "シャフト", 2, "S45C", "東和金属", 21), ("B-301", "ベアリング 6204", 4, "-", "日昇ベアリング", 14),
             ("H-401", "ハウジング", 1, "FC250", "丸栄鋳造", 35), ("C-501", "カバー", 1, "ADC12", "丸栄鋳造", 30),
             ("O-601", "オイルシール", 2, "NBR", "協和パッキン", 10), ("K-701", "キー", 2, "S45C", "社内加工", 3)]
    for i, x in enumerate(parts):
        ws.write_row(3 + i, 0, x, F["c"])
    ws.write_comment("E6", "2026/09出荷分より北辰鋼業へ切替。切替後のリードタイムは28日。（購買課 鈴木）",
                     {"x_scale": 2, "y_scale": 1.5})
    ws.set_column(0, 5, 15)
    finish(wb, p)


def stock():
    wb, p = wb_new("在庫表_2026-09末"); F = fmts(wb)
    ws = wb.add_worksheet("在庫")
    ws.write(0, 0, "部品在庫一覧（2026年9月30日 棚卸）", F["t"])
    ws.write_row(2, 0, ["部品番号", "部品名", "在庫数", "安全在庫", "保管場所"], F["h"])
    rows = [("G-101", "入力ギア", 85, 50, "A-1"), ("G-102", "出力ギア", 92, 50, "A-1"), ("S-201", "シャフト", 40, 60, "A-3"),
            ("B-301", "ベアリング 6204", 310, 200, "B-2"), ("H-401", "ハウジング", 55, 30, "C-1"), ("C-501", "カバー", 48, 30, "C-1"),
            ("O-601", "オイルシール", 140, 100, "B-4"), ("K-701", "キー", 400, 100, "B-5"), ("P-501", "プレート素材", 1200, 800, "D-1")]
    for i, x in enumerate(rows):
        ws.write_row(3 + i, 0, x, F["c"])
    ws.set_column(0, 4, 14)
    finish(wb, p)


# ---------------------------------------------------------------- heat treatment flow (shapes + connectors)
def cxn(cid, name, a, b, x, y, cx, cy, flipv=False):
    fv = ' flipV="1"' if flipv else ""
    return (f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{cid}" name="{name}"/><xdr:cNvCxnSpPr>'
            f'<a:stCxn id="{a}" idx="2"/><a:endCxn id="{b}" idx="0"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr>'
            f'<xdr:spPr><a:xfrm{fv}><a:off x="{x}" y="{y}"/><a:ext cx="{max(cx,1)}" cy="{max(cy,1)}"/></a:xfrm>'
            f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill>'
            f'<a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></xdr:spPr></xdr:cxnSp>')


def heat_flow():
    wb, p = wb_new("作業手順書_熱処理"); F = fmts(wb)
    ws = wb.add_worksheet("手順")
    ws.write(0, 0, "作業手順書：歯車部品の熱処理（GX-200 入力ギア/出力ギア）", F["t"])
    ws.write(1, 0, "文書番号 WI-HT-003　改訂2　2026/06/01")
    ws.write(2, 0, "※硬度の合否基準は検査基準書（旧_検査基準.xls）による。")
    W, H = 150 * EMU, 50 * EMU
    # node: id, text, col, row, prst
    nodes = {11: ("受入検査", 1, 5, "flowChartProcess"), 12: ("洗浄・脱脂", 1, 9, "flowChartProcess"),
             13: ("焼入れ", 1, 13, "flowChartProcess"), 14: ("焼戻し", 1, 17, "flowChartProcess"),
             15: ("硬度判定", 1, 21, "flowChartDecision"), 16: ("出荷（合格時）", 1, 26, "flowChartTerminator"),
             17: ("再処理（不合格時）", 5, 21, "flowChartProcess")}
    order = list(nodes); rng.shuffle(order)          # XML order != flow order
    anchors = []
    for nid in order:
        t, c, r, pr = nodes[nid]
        x, y = B.pos(c, r)
        anchors.append(B.anchor(c, r, c + 3, r + 3, B.sp(nid, f"図形 {nid}", t, x, y, W, H, pr)))
    edges = [(11, 12), (12, 13), (13, 14), (14, 15), (15, 16), (15, 17), (17, 14)]
    rng.shuffle(edges)
    for k, (a, b) in enumerate(edges):
        ca, ra = nodes[a][1], nodes[a][2]; cb, rb = nodes[b][1], nodes[b][2]
        x1, y1 = B.pos(ca + 1, ra + 3); x2, y2 = B.pos(cb + 1, rb)
        anchors.append(B.anchor(min(ca, cb) + 1, min(ra, rb) + 3, max(ca, cb) + 2, max(ra, rb) + 1,
                                cxn(40 + k, f"直線矢印コネクタ {40+k}", a, b, min(x1, x2), min(y1, y2),
                                    abs(x2 - x1), abs(y2 - y1), flipv=y2 < y1)))
    x, y = B.pos(6, 5)
    anchors.append(B.anchor(6, 5, 11, 10, B.sp(60, "テキスト ボックス 60",
                    [B._para("焼入れ条件", "l"), B._para("加熱温度：850℃±10℃", "l"), B._para("保持時間：30分", "l"),
                     B._para("冷却：油冷（油温60〜80℃）", "l")], x, y, 5 * 64 * EMU, 5 * 20 * EMU, "rect", fill="FFF2CC")))
    x, y = B.pos(6, 13)
    anchors.append(B.anchor(6, 13, 11, 16, B.sp(61, "テキスト ボックス 61",
                    [B._para("焼戻し条件：180℃×90分、空冷", "l")], x, y, 5 * 64 * EMU, 3 * 20 * EMU, "rect", fill="FFF2CC")))
    B.placeholder(ws, p, 1, 5, "".join(anchors))
    finish(wb, p)


def legacy_inspection():
    wb, p = wb_new("旧_検査基準"); F = fmts(wb)
    ws = wb.add_worksheet("硬度基準")
    ws.write(0, 0, "検査基準書（熱処理品 硬度）　QS-12", F["t"])
    ws.write_row(2, 0, ["対象部品", "測定方法", "基準", "抜取数"], F["h"])
    rows = [("GX-200 入力ギア/出力ギア", "ロックウェルC", "HRC58〜62", "1ロット5個"),
            ("HS-110 シャフト", "ロックウェルC", "HRC40〜45", "1ロット3個"),
            ("BR-30 ブラケット", "（熱処理なし）", "-", "-")]
    for i, x in enumerate(rows):
        ws.write_row(3 + i, 0, x, F["c"])
    ws.set_column(0, 3, 24)
    finish(wb, p)
    return p


# ---------------------------------------------------------------- press inspection + nameplate image
def nameplate(path):
    img = Image.new("RGB", (620, 360), "#B8B8B0"); d = ImageDraw.Draw(img)
    d.rounded_rectangle([20, 20, 600, 340], radius=14, fill="#D9D9D0", outline="#555555", width=4)
    for (x, y) in [(40, 40), (580, 40), (40, 320), (580, 320)]:
        d.ellipse([x - 8, y - 8, x + 8, y + 8], fill="#777777")
    d.text((180, 45), "ミナト精工株式会社", font=IM.f(28, IM.SANS_B), fill="#202020")
    lines = [("名称", "油圧プレス（3号機）"), ("型式", "MP-200S"), ("定格能力", "2000 kN（200 tf）"),
             ("製造番号", "PR-3-0457"), ("製造年月", "2015年 3月")]
    y = 105
    for k, v in lines:
        d.text((70, y), k, font=IM.f(22), fill="#303030"); d.text((240, y), v, font=IM.f(22, IM.SANS_B), fill="#101010"); y += 42
    img = img.rotate(-3, fillcolor="#A0A098", expand=False)
    img.convert("RGB").save(path, "JPEG", quality=70)


def press():
    wb, p = wb_new("設備点検表_プレス3号機"); F = fmts(wb)
    ws = wb.add_worksheet("点検記録")
    ws.write(0, 0, "月次設備点検表　プレス3号機", F["t"])
    ws.write_row(2, 0, ["点検日", "点検者", "油圧(MPa)", "異常有無", "備考"], F["h"])
    recs = [("2026/04/15", "高橋", 14.0, "なし", ""), ("2026/05/15", "高橋", 13.9, "なし", ""),
            ("2026/06/16", "伊藤", 14.1, "なし", "ガード清掃"), ("2026/07/15", "高橋", 13.8, "なし", ""),
            ("2026/08/18", "伊藤", 13.6, "なし", "次回油圧注視"), ("2026/09/15", "高橋", 12.1, "あり", "油圧低下（基準13.5MPa以上）")]
    for i, x in enumerate(recs):
        ws.write_row(3 + i, 0, x, F["c"])
    ws.write_comment("E9", "9/18 シリンダシールを交換。交換後の油圧14.0MPaで再点検OK。（保全 高橋）", {"x_scale": 2})
    ws.set_column(0, 4, 16); ws.set_column(4, 4, 30)
    info = wb.add_worksheet("設備情報")
    info.write(0, 0, "設備仕様は銘板写真を参照。", F["t"])
    info.write(1, 0, "設置場所：第1工場 プレスライン")
    ip = f"{TMP}/nameplate.jpg"; nameplate(ip)
    info.insert_image("A4", ip, {"description": "銘板写真"})
    finish(wb, p)


# ---------------------------------------------------------------- defect report
def defect_sketch(path):
    img = Image.new("RGB", (640, 380), "white"); d = ImageDraw.Draw(img)
    d.rectangle([80, 90, 420, 300], outline="black", width=3)
    d.ellipse([200, 160, 300, 230], outline="black", width=2)
    for x in range(420, 432, 3):
        d.line([x, 120, x + 4, 150], fill="red", width=2)
    d.ellipse([400, 100, 450, 170], outline="red", width=3)
    d.line([450, 135, 520, 80], fill="red", width=2)
    d.text((470, 40), "バリ高さ 0.3mm", font=IM.f(22, IM.SANS_B), fill="red")
    d.text((470, 70), "（規格 0.1mm以下）", font=IM.f(16), fill="red")
    d.text((90, 320), "PL-50 外周部 / 不良品No.3 スケッチ", font=IM.f(18), fill="black")
    img.save(path)


def defect(order_no, qty):
    wb, p = wb_new("不具合報告_2026-031"); F = fmts(wb)
    ws = wb.add_worksheet("報告書")
    ws.write(0, 0, "不具合報告書", F["t"])
    items = [("報告No", "2026-031"), ("発生日", "2026/09/22"), ("品番", "PL-50 プレート"), ("対象受注No", order_no),
             ("不良数", f"{qty*6//25}個 / ロット{qty}個"), ("現象", "外周部にバリ発生（スケッチ参照）"),
             ("原因", "プレス金型の刃先摩耗"), ("恒久対策", "金型修正（2026/10/10 完了予定）"), ("報告者", "品質保証課 中村")]
    for i, (k, v) in enumerate(items):
        ws.write(2 + i, 0, k, F["h"]); ws.write(2 + i, 1, v, F["c"])
    ws.set_column(0, 0, 14); ws.set_column(1, 1, 40)
    x, y = B.pos(4, 2)
    B.placeholder(ws, p, 4, 2, B.anchor(4, 2, 9, 7, B.sp(70, "吹き出し 70",
                  [B._para("暫定対策：全数目視検査＋バリ取り後に出荷", "l")], x, y, 5 * 64 * EMU, 4 * 20 * EMU,
                  "wedgeRectCallout", fill="FCE4D6")))
    ip = f"{TMP}/defect.png"; defect_sketch(ip)
    ws.insert_image("A13", ip, {"description": "図1", "x_scale": 0.8, "y_scale": 0.8})
    finish(wb, p)


# ---------------------------------------------------------------- minutes / training / change / customer / emergency
def minutes():
    wb, p = wb_new("会議議事録_品質会議_0915"); F = fmts(wb)
    ws = wb.add_worksheet("議事録")
    ws.write(0, 0, "品質会議 議事録（2026/09/15 10:00〜11:30 第1会議室）", F["t"])
    ws.write(1, 0, "出席：山本部長、中村、高橋、伊藤、鈴木（購買）")
    ws.write_row(3, 0, ["No", "議題", "決定事項", "担当", "期限"], F["h"])
    rows = [(1, "8月度不良率報告", "不良率0.42%。目標0.5%以下を達成。", "中村", "-"),
            (2, "プレス3号機 油圧推移", "8月に低下傾向。9月点検で重点確認する。", "高橋", "9/15"),
            (3, "S-201 仕入先切替", "北辰鋼業への切替を承認。初回ロットは全数受入検査。", "鈴木", "9/30"),
            (4, "教育計画", "フォークリフト講習未受講者を11月までに受講させる。", "伊藤", "11/30")]
    for i, x in enumerate(rows):
        ws.write_row(4 + i, 0, x, F["w"])
    ws.merge_range(10, 0, 11, 4, "次回：2026年10月20日(火) 14:00〜　第2会議室", wb.add_format({"bold": True, "border": 1}))
    ws.set_column(1, 2, 34); ws.set_column(3, 4, 8)
    finish(wb, p)


def training():
    wb, p = wb_new("教育記録_2026"); F = fmts(wb)
    ws = wb.add_worksheet("受講状況")
    ws.write(0, 0, "資格・講習 受講状況（2026年度）", F["t"])
    ws.write_row(2, 0, ["社員No", "氏名", "所属", "フォークリフト", "玉掛け", "有機溶剤"], F["h"])
    sur = ["田中", "鈴木", "佐藤", "高橋", "伊藤", "渡辺", "山本", "中村", "小林", "加藤", "吉田", "山田", "佐々木", "山口", "松本",
           "井上", "木村", "林", "清水", "山崎", "森", "池田", "橋本", "阿部", "石川", "前田", "藤田", "小川", "岡田", "後藤"]
    given = ["一郎", "健", "誠", "翔", "大輔", "直樹", "亮", "拓也", "恵", "美咲"]
    fork = 0
    for i, s in enumerate(sur):
        dept = rng.choice(["製造1課", "製造2課", "品質保証課", "保全課", "物流課"])
        f = rng.random() < 0.55
        fork += f
        ws.write_row(3 + i, 0, [f"E{1001+i}", f"{s} {rng.choice(given)}", dept,
                                f"2026/{rng.randint(4,9):02d}/{rng.randint(1,28):02d}" if f else "",
                                "済" if rng.random() < 0.4 else "", "済" if rng.random() < 0.3 else ""], F["c"])
    ws.write(34, 0, "※フォークリフト欄は受講日。空欄は未受講。")
    ws.set_column(0, 5, 13)
    finish(wb, p)
    return fork


def changes():
    wb, p = wb_new("変更管理表"); F = fmts(wb)
    ws = wb.add_worksheet("変更一覧")
    ws.write_row(0, 0, ["変更No", "対象", "変更内容", "申請日", "承認日", "承認者"], F["h"])
    base = [("CH-110", "BR-30", "穴位置公差を±0.1→±0.05に変更", "2026/04/08", "2026/04/15", "山本部長"),
            ("CH-111", "HS-110", "梱包仕様を段ボール→通い箱に変更", "2026/04/20", "2026/04/24", "山本部長"),
            ("CH-114", "GX-200", "グリース銘柄変更", "2026/05/30", "2026/06/05", "山本部長"),
            ("CH-116", "WI-HT-003", "熱処理手順書 改訂2（油温管理を追加）", "2026/05/25", "2026/06/01", "山本部長"),
            ("CH-118", "PL-50", "プレス金型の材質を SKD11 から SKH51 に変更（摩耗対策）", "2026/08/20", "2026/08/28", "山本部長"),
            ("CH-120", "S-201", "仕入先を東和金属→北辰鋼業に変更", "2026/08/25", "2026/09/16", "山本部長"),
            ("CH-121", "GX-200", "銘板ラベルのデザイン変更", "2026/09/10", "", "")]
    for i, x in enumerate(base):
        ws.write_row(1 + i, 0, x, F["c"])
    ws.set_column(0, 1, 10); ws.set_column(2, 2, 52); ws.set_column(3, 5, 12)
    finish(wb, p)


def customer_terms():
    wb, p = wb_new("顧客別納入条件"); F = fmts(wb)
    ws = wb.add_worksheet("納入条件")
    ws.write_row(0, 0, ["顧客", "納入場所", "納入条件", "支払条件"], F["h"])
    rows = [("A社", "本社工場 第2倉庫", "通常便（指定なし）", "月末締め翌月末払い"),
            ("B工業", "浜松工場", "前日15時までに出荷連絡", "月末締め翌々月10日払い"),
            ("C製作所", "本社 受入口", "午前中必着。検査成績書を同梱。指定ラベル（CS-L2）を外箱に貼付。", "20日締め翌月末払い"),
            ("D電機", "各務原工場", "パレット積み（1段あたり最大20箱）", "月末締め翌月末払い"),
            ("E産業", "本社", "納品書2部（控え返送）", "月末締め翌月20日払い")]
    for i, x in enumerate(rows):
        ws.write_row(1 + i, 0, x, F["w"])
    ws.set_column(0, 1, 16); ws.set_column(2, 2, 50); ws.set_column(3, 3, 24)
    finish(wb, p)


def emergency():
    wb, p = wb_new("緊急連絡体制"); F = fmts(wb)
    ws = wb.add_worksheet("連絡体制図")
    ws.write(0, 0, "緊急時連絡体制（2026年度）", F["t"])
    W, H = 200 * EMU, 55 * EMU
    boxes = [(81, "工場長 山本（内線7001）", 4, 3), (82, "平日日中の設備故障：保全課（内線7200）", 1, 9),
             (83, "夜間・休日の設備故障：守衛室（内線7110）→ 保全当番へ連絡", 6, 9),
             (84, "火災・けが：119番通報後、総務課（内線7300）", 11, 9)]
    xml = []
    for sid, t, c, r in boxes:
        x, y = B.pos(c, r)
        xml.append(B.anchor(c, r, c + 4, r + 3, B.sp(sid, f"四角形 {sid}", [B._para(t)], x, y, W, H, "roundRect")))
    B.placeholder(ws, p, 1, 3, "".join(xml))
    finish(wb, p)


def noise(name, kind):
    wb, p = wb_new(name); F = fmts(wb)
    ws = wb.add_worksheet("Sheet1")
    if kind == "daily":
        ws.write_row(0, 0, ["日付", "工程", "作業者", "出来高", "不良", "備考"], F["h"])
        for i in range(20):
            ws.write_row(1 + i, 0, [f"2026/09/{i+1:02d}", rng.choice(["切削", "プレス", "組立", "熱処理"]),
                                    rng.choice(["田中", "佐藤", "高橋", "伊藤"]), rng.randint(80, 300), rng.randint(0, 4),
                                    rng.choice(["", "", "段取り替え", "材料待ち"])], F["c"])
    elif kind == "budget":
        ws.write_row(0, 0, ["費目", "上期予算", "上期実績", "差異"], F["h"])
        for i, k in enumerate(["材料費", "外注費", "消耗品費", "修繕費", "教育費", "旅費"]):
            b = rng.randint(100, 900) * 10000; a = int(b * rng.uniform(0.8, 1.15))
            ws.write_row(1 + i, 0, [k, b, a], F["n"]); ws.write_formula(1 + i, 3, f"=C{i+2}-B{i+2}", F["n"], a - b)
    else:
        ws.write_row(0, 0, ["No", "提案内容", "提案者", "効果(千円/年)"], F["h"])
        for i, t in enumerate(["工具棚の配置見直し", "照明のLED化", "段取り替え治具の共通化", "梱包材の再利用"]):
            ws.write_row(1 + i, 0, [i + 1, t, rng.choice(["田中", "小林", "加藤"]), rng.randint(50, 900)], F["c"])
    ws.set_column(0, 5, 14)
    finish(wb, p)


def main():
    for d in (OUT, TMP):
        shutil.rmtree(d, ignore_errors=True); os.makedirs(d)
    B.PENDING.clear()
    up = orders("受注一覧_2026上期", (2026, 4), 6, "J-26")
    orders("受注一覧_2025下期", (2025, 10), 6, "J-25")
    approval("職務権限規程_決裁権限一覧", 5, "2026/04/01", 300, True)
    approval("職務権限規程_決裁権限一覧_旧版", 4, "2024/04/01", 200, False)
    bom(); stock(); heat_flow()
    xls_src = legacy_inspection()
    press()
    # pick a PL-50 order in Sep (not cancelled) for the defect report
    cand = [r for r in up if r["pn"] == "PL-50" and r["date"].month == 9 and r["st"] != "キャンセル" and r["date"].day < 22 and r["qty"] >= 50]
    tgt = cand[0]
    defect(tgt["no"], tgt["qty"])
    minutes(); fork = training(); changes(); customer_terms(); emergency()
    noise("日報_2026-09_製造1課", "daily"); noise("日報_2026-09_製造2課", "daily")
    noise("経費予算_2026上期", "budget"); noise("改善提案一覧_2026", "kaizen")
    # .xls conversion
    subprocess.run(["soffice", "--headless", "--convert-to", "xls", "--outdir", OUT, xls_src],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    os.remove(xls_src)
    shutil.rmtree(TMP)

    # ---------------- answers
    valid = [r for r in up if r["st"] != "キャンセル"]
    b_q1 = sum(r["amt"] for r in valid if r["cust"] == "B工業" and 4 <= r["date"].month <= 6)
    b_q1_incl = sum(r["amt"] for r in up if r["cust"] == "B工業" and 4 <= r["date"].month <= 6)
    tot = {}
    for r in valid:
        tot[r["cust"]] = tot.get(r["cust"], 0) + r["amt"]
    top = max(tot, key=tot.get)
    Q = [
        ("Q01", "設備を250万円（税抜）で購入したい。決裁者は誰？", "部長（第5版 2026/04/01施行で部長上限が300万円に。旧版第4版では工場長）", "版の新旧"),
        ("Q02", f"B工業の2026年4〜6月の受注金額合計（税抜）は？", f"{b_q1:,}円（キャンセル除く。キャンセル込みだと{b_q1_incl:,}円）", "集計"),
        ("Q03", "2026年上期（4〜9月）で受注金額が最も大きい顧客とその金額は？", f"{top} {tot[top]:,}円（キャンセル除く）", "集計"),
        ("Q04", "GX-200のシャフトの仕入先とリードタイムは？", "北辰鋼業、28日（2026/09出荷分から。表の東和金属/21日は旧情報。コメント・変更管理CH-120）", "コメント"),
        ("Q05", "GX-200のシャフトの在庫は安全在庫を満たしている？", "満たしていない。在庫40に対し安全在庫60（20不足）", "ファイル横断"),
        ("Q06", "熱処理工程で、焼入れの次の工程は？", "焼戻し", "図形のつながり"),
        ("Q07", "硬度判定で不合格になったらどうなる？", "再処理を行い、焼戻し工程に戻る", "図形のつながり"),
        ("Q08", "焼入れの加熱温度・保持時間・冷却方法は？", "850℃±10℃、30分、油冷（油温60〜80℃）", "テキストボックス"),
        ("Q09", "GX-200のギアの熱処理後の硬度基準は？", "HRC58〜62（ロックウェルC、1ロット5個抜取）", ".xls"),
        ("Q10", "プレス3号機の定格能力と製造番号は？", "2000kN（200tf）、PR-3-0457", "画像"),
        ("Q11", "プレス3号機の9月点検で見つかった異常と、その処置は？", "油圧低下（12.1MPa、基準13.5以上）。9/18にシリンダシール交換、14.0MPaで再点検OK", "コメント"),
        ("Q12", "不具合報告2026-031の暫定対策と恒久対策は？", "暫定：全数目視検査＋バリ取り後に出荷／恒久：金型修正（2026/10/10完了予定）", "図形"),
        ("Q13", "不具合報告2026-031のバリの高さと規格値は？", "バリ高さ0.3mm、規格0.1mm以下", "画像"),
        ("Q14", "不具合報告2026-031の対象ロットの顧客と受注数量は？", f"{tgt['cust']}、{tgt['qty']}個（受注No {tgt['no']}）", "ファイル横断"),
        ("Q15", "次回の品質会議の日時と場所は？", "2026年10月20日(火) 14:00〜、第2会議室", "結合セル"),
        ("Q16", "フォークリフト講習を受講済みの人は何人？", f"{fork}人（30人中）", "集計"),
        ("Q17", "変更No CH-118 の内容と承認日は？", "PL-50プレス金型の材質をSKD11→SKH51に変更（摩耗対策）。承認日2026/08/28（山本部長）", "検索"),
        ("Q18", "C製作所に納品するときの注意点は？", "午前中必着、検査成績書を同梱、指定ラベル（CS-L2）を外箱に貼付", "検索"),
        ("Q19", "夜間にプレス機が故障したら、どこに連絡する？", "守衛室（内線7110）に連絡し、保全当番へつなぐ", "図形"),
        ("Q20", "プレス3号機の購入金額は？", "記載なし（どの文書にもない）", "回答不能"),
        ("Q21", "GX-300の部品構成を教えて", "記載なし（GX-300の部品表は無い）", "回答不能"),
        ("Q22", "S-201の仕入先切替は会議で承認されている？初回ロットの扱いは？", "9/15品質会議で承認（変更管理CH-120は9/16承認）。初回ロットは全数受入検査", "ファイル横断"),
    ]
    json.dump([dict(id=a, q=b, answer=c, type=d) for a, b, c, d in Q], open(KEY, "w"), ensure_ascii=False, indent=1)
    print(sorted(os.listdir(OUT)))
    for a, b, c, d in Q:
        print(a, d, "|", b, "=>", c)


if __name__ == "__main__":
    main()
