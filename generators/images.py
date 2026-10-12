"""Test image generator (PNG/JPEG with Japanese text rendered as pixels)."""
import random
from PIL import Image, ImageDraw, ImageFont, ImageFilter

SANS = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
SANS_B = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
SERIF = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"


def f(size, path=SANS):
    # index 0 of the .ttc is JP in Noto CJK collections
    return ImageFont.truetype(path, size, index=0)


def _save(img, path, jpeg=False, quality=90):
    if jpeg:
        img.convert("RGB").save(path, "JPEG", quality=quality)
    else:
        img.save(path, "PNG")
    return path


def clear(text, path):
    img = Image.new("RGB", (520, 200), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 519, 44], fill="#1F4E79")
    d.text((14, 6), "工程連絡票", font=f(24, SANS_B), fill="white")
    d.text((20, 70), "区分：", font=f(26), fill="black")
    d.text((110, 70), text, font=f(26, SANS_B), fill="black")
    d.text((20, 130), "担当：品質保証課 / 期日：11月末", font=f(18), fill="#333333")
    return _save(img, path)


def small(text, path):
    img = Image.new("RGB", (640, 220), "white")
    d = ImageDraw.Draw(img)
    y = 10
    lines = ["・治具の摩耗を点検すること", "・トルク値は規定範囲内を確認",
             f"・最終判定は{text}のうえ出荷", "・記録は3年間保管", "・異常時は直ちにライン停止"]
    for ln in lines:
        d.text((10, y), ln, font=f(11), fill="black")
        y += 18
    d.rectangle([380, 20, 620, 200], outline="#999999")
    d.text((400, 100), "（図面参照）", font=f(16), fill="#777777")
    return _save(img, path)


def rotated(text, path):
    base = Image.new("RGBA", (420, 120), (255, 255, 255, 0))
    d = ImageDraw.Draw(base)
    d.rectangle([0, 0, 419, 119], outline="#C00000", width=4)
    d.text((24, 32), text, font=f(40, SANS_B), fill="#C00000")
    rot = base.rotate(23, expand=True, resample=Image.BICUBIC)
    img = Image.new("RGB", (560, 360), "white")
    img.paste(rot, (40, 20), rot)
    return _save(img, path)


def lowcontrast(text, path):
    img = Image.new("RGB", (520, 160), "#D9D9D9")
    d = ImageDraw.Draw(img)
    d.text((30, 50), f"備考：{text}は課長決裁", font=f(28), fill="#C4C4C4")
    return _save(img, path)


def noisy(text, path):
    img = Image.new("RGB", (560, 200), "#F2F2F2")
    d = ImageDraw.Draw(img)
    d.text((24, 30), "【スキャン】設備変更申請", font=f(26, SANS_B), fill="#202020")
    d.text((24, 100), f"判定欄：{text}", font=f(24), fill="#303030")
    rng = random.Random(7)
    px = img.load()
    for _ in range(9000):
        x, y = rng.randrange(560), rng.randrange(200)
        v = rng.randrange(60, 230)
        px[x, y] = (v, v, v)
    img = img.filter(ImageFilter.GaussianBlur(1.1)).rotate(1.5, fillcolor="#F2F2F2")
    return _save(img, path, jpeg=True, quality=35)


def vertical(text, path):
    img = Image.new("RGB", (200, 420), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([40, 10, 160, 410], outline="black", width=2)
    y = 30
    for ch in text:
        d.text((78, y), ch, font=f(40, SERIF), fill="black")
        y += 52
    return _save(img, path)


def stamp(text, path):
    img = Image.new("RGB", (360, 360), "white")
    d = ImageDraw.Draw(img)
    d.ellipse([30, 30, 330, 330], outline="#D00000", width=8)
    d.line([40, 140, 320, 140], fill="#D00000", width=4)
    d.line([40, 220, 320, 220], fill="#D00000", width=4)
    d.text((95, 80), "2026.10.12", font=f(30, SANS_B), fill="#D00000")
    d.text((62, 155), text, font=f(40, SANS_B), fill="#D00000")
    d.text((140, 235), "品証", font=f(36, SANS_B), fill="#D00000")
    return _save(img, path)


def diagram(text, path):
    img = Image.new("RGB", (900, 260), "white")
    d = ImageDraw.Draw(img)
    boxes = ["受入検査", "加工", "組立", text, "出荷"]
    x = 20
    for i, b in enumerate(boxes):
        d.rounded_rectangle([x, 90, x + 150, 170], radius=12, outline="#2F5597", width=3, fill="#DDEBF7")
        d.text((x + 12, 115), b, font=f(20), fill="black")
        if i < len(boxes) - 1:
            d.line([x + 150, 130, x + 175, 130], fill="black", width=3)
            d.polygon([(x + 175, 122), (x + 175, 138), (x + 185, 130)], fill="black")
        x += 185
    d.text((20, 20), "製造フロー（改訂3）", font=f(26, SANS_B), fill="black")
    return _save(img, path)


def screenshot(text, path):
    img = Image.new("RGB", (700, 300), "#FFFFFF")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 699, 30], fill="#E1E1E1")
    d.text((10, 4), "生産管理システム - 案件一覧", font=f(15), fill="black")
    for i, c in enumerate(["#ECECEC", "#3C78D8"]):
        pass
    rows = [("案件No", "品名", "状態"), ("A-1021", "ブラケット", "仕掛"),
            ("A-1022", "シャフト", text), ("A-1023", "カバー", "完了"), ("A-1024", "ギア", "保留")]
    y = 50
    for r, row in enumerate(rows):
        if r == 0:
            d.rectangle([10, y - 4, 690, y + 26], fill="#DDEBF7")
        xs = [20, 180, 400]
        for x, cell in zip(xs, row):
            d.text((x, y), cell, font=f(17), fill="black")
        d.line([10, y + 28, 690, y + 28], fill="#BBBBBB")
        y += 44
    return _save(img, path)


def serif_skew(text, path):
    base = Image.new("RGB", (560, 160), "white")
    d = ImageDraw.Draw(base)
    d.text((30, 40), f"※{text}待ち", font=f(44, SERIF), fill="#1A1A66")
    img = base.transform(base.size, Image.AFFINE, (1, 0.35, -40, 0, 1, 0), fillcolor="white")
    return _save(img, path)


PIXEL_STYLES = {
    "clear": clear, "small": small, "rotated": rotated, "lowcontrast": lowcontrast,
    "noisy": noisy, "vertical": vertical, "stamp": stamp, "diagram": diagram,
    "screenshot": screenshot, "serif_skew": serif_skew,
}

DECOY_TEXTS = ["確認済", "要再検査", "部長決裁", "暫定対応", "出荷可", "保留中", "完了", "差戻し"]


def decoy(path, rng):
    style = rng.choice(["clear", "diagram", "screenshot", "stamp"])
    return PIXEL_STYLES[style](rng.choice(DECOY_TEXTS), path)


def photo_like(path, rng):
    """Image with no text at all (for alt-text tests)."""
    img = Image.new("RGB", (400, 260), "#87A96B")
    d = ImageDraw.Draw(img)
    for _ in range(40):
        x, y = rng.randrange(400), rng.randrange(260)
        r = rng.randrange(10, 60)
        c = tuple(rng.randrange(80, 220) for _ in range(3))
        d.ellipse([x - r, y - r, x + r, y + r], fill=c)
    return _save(img, path)
