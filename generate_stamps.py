"""
LINEスタンプ画像生成スクリプト
LINE Creators Market 仕様:
  - メインスタンプ: 370x320 px, PNG (透過)
  - タブ画像:      96x74 px,  PNG (透過)
  - セット枚数:    8枚（最小単位）
"""

from PIL import Image, ImageDraw, ImageFont
import os
import math

OUT_MAIN = "stamps/main"
OUT_TAB  = "stamps/tab"
W, H     = 370, 320   # メインスタンプサイズ
TW, TH   = 96, 74     # タブ画像サイズ

# カラーパレット
BODY_COLOR    = (255, 220, 100, 255)   # 黄色い体
OUTLINE_COLOR = (60, 40, 10, 255)      # 濃いアウトライン
CHEEK_COLOR   = (255, 160, 140, 180)   # ほっぺ
WHITE         = (255, 255, 255, 255)
BLACK         = (0, 0, 0, 255)
RED           = (220, 50, 50, 255)
BLUE          = (80, 130, 220, 255)
PINK          = (255, 100, 160, 255)
GREEN         = (80, 190, 100, 255)


def new_canvas():
    return Image.new("RGBA", (W, H), (0, 0, 0, 0))


def draw_face(draw, cx, cy, r, eye_style="normal", mouth_style="smile",
              blush=True, extra_fn=None):
    """キャラクターの顔を描画する共通関数"""
    # 体（円）
    draw.ellipse([cx-r, cy-r, cx+r, cy+r],
                 fill=BODY_COLOR, outline=OUTLINE_COLOR, width=4)

    # ほっぺ
    if blush:
        cr = r // 5
        for bx in [cx - int(r * 0.52), cx + int(r * 0.52)]:
            draw.ellipse([bx-cr, cy+cr//2-cr, bx+cr, cy+cr//2+cr],
                         fill=CHEEK_COLOR)

    # 目
    ew = max(6, r // 7)
    eh = max(8, r // 6)
    ex1, ex2 = cx - int(r * 0.32), cx + int(r * 0.32)
    ey = cy - int(r * 0.12)

    if eye_style == "normal":
        draw.ellipse([ex1-ew, ey-eh, ex1+ew, ey+eh], fill=BLACK)
        draw.ellipse([ex2-ew, ey-eh, ex2+ew, ey+eh], fill=BLACK)
        # ハイライト
        for ex in [ex1, ex2]:
            hw = max(2, ew // 2)
            draw.ellipse([ex-hw+2, ey-eh+2, ex-hw+2+hw, ey-eh+2+hw], fill=WHITE)

    elif eye_style == "happy":
        # ^^目
        for ex in [ex1, ex2]:
            draw.arc([ex-ew, ey-eh//2, ex+ew, ey+eh//2],
                     start=200, end=340, fill=BLACK, width=4)

    elif eye_style == "closed":
        for ex in [ex1, ex2]:
            draw.line([ex-ew, ey, ex+ew, ey], fill=BLACK, width=4)

    elif eye_style == "surprise":
        draw.ellipse([ex1-ew*2, ey-eh*2, ex1+ew*2, ey+eh*2], fill=WHITE, outline=BLACK, width=3)
        draw.ellipse([ex1-ew, ey-eh, ex1+ew, ey+eh], fill=BLACK)
        draw.ellipse([ex2-ew*2, ey-eh*2, ex2+ew*2, ey+eh*2], fill=WHITE, outline=BLACK, width=3)
        draw.ellipse([ex2-ew, ey-eh, ex2+ew, ey+eh], fill=BLACK)

    elif eye_style == "angry":
        for ex in [ex1, ex2]:
            draw.ellipse([ex-ew, ey-eh, ex+ew, ey+eh], fill=BLACK)
        # 怒り眉
        bw = int(r * 0.18)
        by = ey - eh - int(r * 0.1)
        draw.line([ex1-bw, by-6, ex1+bw, by+4], fill=OUTLINE_COLOR, width=5)
        draw.line([ex2-bw, by+4, ex2+bw, by-6], fill=OUTLINE_COLOR, width=5)

    elif eye_style == "tear":
        draw.ellipse([ex1-ew, ey-eh, ex1+ew, ey+eh], fill=BLACK)
        draw.ellipse([ex2-ew, ey-eh, ex2+ew, ey+eh], fill=BLACK)
        # 涙
        for ex in [ex1, ex2]:
            ty = ey + eh
            draw.ellipse([ex-4, ty, ex+4, ty+int(r*0.35)], fill=BLUE)

    # 口
    mx, my = cx, cy + int(r * 0.32)
    mw = int(r * 0.38)

    if mouth_style == "smile":
        draw.arc([mx-mw, my-mw//2, mx+mw, my+mw//2],
                 start=20, end=160, fill=OUTLINE_COLOR, width=4)

    elif mouth_style == "big_smile":
        draw.arc([mx-mw, my-mw//2, mx+mw, my+mw//2],
                 start=10, end=170, fill=OUTLINE_COLOR, width=5)
        draw.arc([mx-mw, my-mw//2, mx+mw, my+mw//2+4],
                 start=10, end=170, fill=WHITE, width=3)

    elif mouth_style == "sad":
        draw.arc([mx-mw, my, mx+mw, my+mw//2],
                 start=200, end=340, fill=OUTLINE_COLOR, width=4)

    elif mouth_style == "open":
        draw.ellipse([mx-mw//2, my-mw//3, mx+mw//2, my+mw//3],
                     fill=RED, outline=OUTLINE_COLOR, width=3)

    elif mouth_style == "straight":
        draw.line([mx-mw//2, my, mx+mw//2, my], fill=OUTLINE_COLOR, width=4)

    if extra_fn:
        extra_fn(draw, cx, cy, r)


def add_text(img, text, color=(60, 40, 10, 255), size=28, y_offset=0):
    """スタンプにテキストを追加"""
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", size)
    except Exception:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    tx = (W - tw) // 2
    ty = H - 55 + y_offset
    # 白縁取り
    for dx, dy in [(-2,0),(2,0),(0,-2),(0,2),(-2,-2),(2,-2),(-2,2),(2,2)]:
        draw.text((tx+dx, ty+dy), text, font=font, fill=(255,255,255,220))
    draw.text((tx, ty), text, font=font, fill=color)


def stamp_happy(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)
    draw_face(d, W//2, H//2 - 20, 110, eye_style="happy", mouth_style="big_smile")
    add_text(img, "ありがとう!")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_sad(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)
    draw_face(d, W//2, H//2 - 20, 110, eye_style="tear", mouth_style="sad")
    add_text(img, "ごめんね...")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_ok(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def thumb(draw, cx, cy, r):
        # 親指グッドアイコン
        px, py = cx + int(r * 0.85), cy + int(r * 0.3)
        draw.ellipse([px-22, py-22, px+22, py+22], fill=GREEN, outline=OUTLINE_COLOR, width=3)
        draw.polygon([(px-6, py-14), (px+6, py-14), (px+6, py+10), (px-6, py+10)],
                     fill=WHITE)
        draw.polygon([(px-12, py-4), (px+8, py-12), (px+8, py-4)], fill=WHITE)

    draw_face(d, W//2 - 15, H//2 - 20, 100, eye_style="happy", mouth_style="smile",
              extra_fn=thumb)
    add_text(img, "了解!")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_love(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def hearts(draw, cx, cy, r):
        for hx, hy, hs in [(cx+r+10, cy-r+10, 24), (cx-r-10, cy-r, 18), (cx+r-10, cy-r-20, 16)]:
            _draw_heart(draw, hx, hy, hs, PINK)

    draw_face(d, W//2, H//2 - 10, 110, eye_style="happy", mouth_style="big_smile",
              extra_fn=hearts)
    add_text(img, "好き♡")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_angry(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def veins(draw, cx, cy, r):
        vx, vy = cx + int(r * 0.7), cy - int(r * 0.8)
        draw.line([vx, vy, vx+12, vy-12, vx+20, vy], fill=RED, width=4)
        draw.line([vx, vy, vx+12, vy+12, vx+20, vy], fill=RED, width=4)

    draw_face(d, W//2, H//2 - 20, 110, eye_style="angry", mouth_style="open",
              extra_fn=veins)
    add_text(img, "ムキー!!", color=RED)
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_surprise(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def sweat(draw, cx, cy, r):
        sx, sy = cx + int(r * 0.9), cy - int(r * 0.5)
        draw.ellipse([sx-6, sy, sx+6, sy+20], fill=BLUE)
        draw.polygon([(sx, sy-8), (sx-6, sy+4), (sx+6, sy+4)], fill=BLUE)

    draw_face(d, W//2, H//2 - 20, 110, eye_style="surprise", mouth_style="open",
              extra_fn=sweat)
    add_text(img, "えっ!?")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_hello(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def wave(draw, cx, cy, r):
        # 手を振る
        ax, ay = cx + int(r * 0.9), cy - int(r * 0.1)
        for i in range(3):
            angle = math.radians(40 + i * 25)
            x2 = ax + int(40 * math.cos(angle))
            y2 = ay - int(40 * math.sin(angle))
            draw.line([ax, ay, x2, y2], fill=BODY_COLOR, width=18)
            draw.line([ax, ay, x2, y2], fill=OUTLINE_COLOR, width=4)
        draw.ellipse([ax-16, ay-16, ax+16, ay+16], fill=BODY_COLOR, outline=OUTLINE_COLOR, width=4)

    draw_face(d, W//2 - 20, H//2 - 20, 100, eye_style="happy", mouth_style="smile",
              extra_fn=wave)
    add_text(img, "こんにちは!")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def stamp_sleep(n):
    img = new_canvas()
    d = ImageDraw.Draw(img)

    def zzz(draw, cx, cy, r):
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
        except Exception:
            font = ImageFont.load_default()
        for i, (zx, zy, sz) in enumerate([(cx+r+5, cy-r+20, 20),
                                            (cx+r+22, cy-r+5, 26),
                                            (cx+r+42, cy-r-12, 32)]):
            try:
                zf = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", sz)
            except Exception:
                zf = font
            draw.text((zx, zy), "z", font=zf, fill=BLUE)

    draw_face(d, W//2, H//2 - 15, 110, eye_style="closed", mouth_style="straight",
              blush=True, extra_fn=zzz)
    add_text(img, "おやすみ~")
    img.save(f"{OUT_MAIN}/{n:02d}.png")


def _draw_heart(draw, cx, cy, size, color):
    """ハートの描画"""
    points = []
    for i in range(0, 360, 5):
        t = math.radians(i)
        x = size * (16 * math.sin(t)**3) / 16
        y = -size * (13*math.cos(t) - 5*math.cos(2*t) - 2*math.cos(3*t) - math.cos(4*t)) / 16
        points.append((cx + x, cy + y))
    if len(points) >= 3:
        draw.polygon(points, fill=color)


def make_tab():
    img = Image.new("RGBA", (TW, TH), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = 28
    cx, cy = TW//2, TH//2 - 2
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=BODY_COLOR, outline=OUTLINE_COLOR, width=3)
    ew, eh = 4, 5
    for ex in [cx-9, cx+9]:
        d.ellipse([ex-ew, cy-eh-2, ex+ew, cy+eh-2], fill=BLACK)
    mw = 12
    d.arc([cx-mw, cy+4, cx+mw, cy+4+mw//2], start=20, end=160, fill=OUTLINE_COLOR, width=3)
    img.save(f"{OUT_TAB}/tab.png")
    print("Saved: stamps/tab/tab.png")


STAMPS = [
    stamp_happy,
    stamp_sad,
    stamp_ok,
    stamp_love,
    stamp_angry,
    stamp_surprise,
    stamp_hello,
    stamp_sleep,
]

if __name__ == "__main__":
    os.makedirs(OUT_MAIN, exist_ok=True)
    os.makedirs(OUT_TAB, exist_ok=True)

    for i, fn in enumerate(STAMPS, start=1):
        fn(i)
        print(f"Saved: stamps/main/{i:02d}.png  ({fn.__name__})")

    make_tab()
    print(f"\n完了: {len(STAMPS)}枚のスタンプ + タブ画像を生成しました")
    print("提出先: LINE Creators Market (https://creator.line.me/)")
