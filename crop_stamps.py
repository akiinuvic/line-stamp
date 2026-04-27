"""
LINEスタンプ合成画像 → 個別PNG切り出しスクリプト
使い方: python3 crop_stamps.py stamps_base.png

出力: output/stamps/main/01.png ~ 24.png
      output/stamps/tab/tab.png
"""

import sys
import os
from PIL import Image, ImageDraw

SOURCE = sys.argv[1] if len(sys.argv) > 1 else "stamps_base.png"
OUT_MAIN = "output/stamps/main"
OUT_TAB  = "output/stamps/tab"

COLS, ROWS = 6, 4
TARGET_W, TARGET_H = 370, 320  # LINE Creators Market 仕様

# 数字バッジ（左上の番号）をカットするためのオフセット
# ※画像によって要調整
TRIM_LEFT   = 28
TRIM_TOP    = 32
TRIM_RIGHT  = 4
TRIM_BOTTOM = 4


def make_transparent_bg(img: Image.Image, tolerance: int = 30) -> Image.Image:
    """白背景を透過に変換（フラッドフィル）"""
    img = img.convert("RGBA")
    data = img.load()
    w, h = img.size

    def is_white(pixel, tol):
        r, g, b, a = pixel
        return r > 255 - tol and g > 255 - tol and b > 255 - tol

    # 四隅からフラッドフィルで白を透過に
    stack = []
    corners = [(0, 0), (w-1, 0), (0, h-1), (w-1, h-1)]
    visited = [[False]*h for _ in range(w)]

    for cx, cy in corners:
        if is_white(data[cx, cy], tolerance):
            stack.append((cx, cy))

    while stack:
        x, y = stack.pop()
        if x < 0 or x >= w or y < 0 or y >= h:
            continue
        if visited[x][y]:
            continue
        visited[x][y] = True
        if is_white(data[x, y], tolerance):
            data[x, y] = (255, 255, 255, 0)
            stack.extend([(x+1, y), (x-1, y), (x, y+1), (x, y-1)])

    return img


def crop_and_resize(cell: Image.Image) -> Image.Image:
    """セルを LINE 仕様サイズにリサイズ（アスペクト比維持・センタリング）"""
    cw, ch = cell.size
    ratio = min(TARGET_W / cw, TARGET_H / ch)
    new_w = int(cw * ratio)
    new_h = int(ch * ratio)
    resized = cell.resize((new_w, new_h), Image.LANCZOS)

    canvas = Image.new("RGBA", (TARGET_W, TARGET_H), (255, 255, 255, 0))
    px = (TARGET_W - new_w) // 2
    py = (TARGET_H - new_h) // 2
    canvas.paste(resized, (px, py), resized)
    return canvas


def make_tab(first_stamp_path: str):
    """タブ画像（96×74）を1枚目のスタンプから生成"""
    img = Image.open(first_stamp_path).convert("RGBA")
    tab = img.resize((96, 74), Image.LANCZOS)
    tab.save(f"{OUT_TAB}/tab.png")
    print("Saved: output/stamps/tab/tab.png")


def main():
    if not os.path.exists(SOURCE):
        print(f"エラー: {SOURCE} が見つかりません")
        print("使い方: python3 crop_stamps.py stamps_base.png")
        sys.exit(1)

    os.makedirs(OUT_MAIN, exist_ok=True)
    os.makedirs(OUT_TAB, exist_ok=True)

    src = Image.open(SOURCE).convert("RGBA")
    W, H = src.size
    print(f"ソース画像: {W}×{H}px")

    cw = W // COLS
    ch = H // ROWS
    print(f"セルサイズ: {cw}×{ch}px (トリム前)")

    for row in range(ROWS):
        for col in range(COLS):
            n = row * COLS + col + 1

            x1 = col * cw + TRIM_LEFT
            y1 = row * ch + TRIM_TOP
            x2 = (col + 1) * cw - TRIM_RIGHT
            y2 = (row + 1) * ch - TRIM_BOTTOM

            cell = src.crop((x1, y1, x2, y2))
            cell = make_transparent_bg(cell)
            cell = crop_and_resize(cell)

            out_path = f"{OUT_MAIN}/{n:02d}.png"
            cell.save(out_path)
            print(f"  [{n:02d}] 保存: {out_path}")

    make_tab(f"{OUT_MAIN}/01.png")
    print(f"\n完了! output/stamps/ に {ROWS*COLS}枚 + タブ画像を生成しました")
    print("次のステップ: LINE Creators Market にアップロード")
    print("  https://creator.line.me/")


if __name__ == "__main__":
    main()
