# -*- coding: utf-8 -*-
"""APP 图标生成：源图 -> 全套 mipmap（legacy / round / adaptive foreground）
用法：python _make_icon.py <源图路径>
规则（v2.39 起）：完整显示源图，不裁剪不修改构图——
  1. 仅检测纯白外围边（min(RGB)>235）裁掉，暗色图不受影响
  2. legacy = 完整方图原样（启动器自行处理形状）
  3. round = 完整图等比缩小置于圆形内（内容完整，圆外透明）
  4. adaptive foreground = 完整图缩至安全区（66/108）居中，四周透出背景色
  5. 背景色取源图四角 20px 均值
"""
import sys, os
from PIL import Image, ImageDraw

RES = r'C:\Users\Administrator\WorkBuddy\2026-10-12-00-40-20\app\android\app\src\main\res'

# 密度 -> (launcher 尺寸, foreground 画布尺寸 108dp)
SIZES = {
    'mdpi': (48, 108), 'hdpi': (72, 162), 'xhdpi': (96, 216),
    'xxhdpi': (144, 324), 'xxxhdpi': (192, 432),
}
SAFE = 66 / 108  # 自适应图标安全区比例（启动器圆形 mask 只显示中间 ~66/108）

def trim_pure_white(img):
    """外围纯白边检测：min(RGB)>235 判白，逐行/列扫描找内容 bbox（暗色图直接返回原图）"""
    w, h = img.size
    px = img.convert('RGB').load()
    def row_white(y):
        return all(min(px[x, y]) > 235 for x in range(0, w, 4))
    def col_white(x):
        return all(min(px[x, y]) > 235 for y in range(0, h, 4))
    top = 0
    while top < h // 3 and row_white(top): top += 1
    bottom = h - 1
    while bottom > 2 * h // 3 and row_white(bottom): bottom -= 1
    left = 0
    while left < w // 3 and col_white(left): left += 1
    right = w - 1
    while right > 2 * w // 3 and col_white(right): right -= 1
    return img.crop((left, top, right + 1, bottom + 1))

def square_center(img):
    w, h = img.size
    s = min(w, h)
    return img.crop(((w - s) // 2, (h - s) // 2, (w + s) // 2, (h + s) // 2))

def corner_bg(img):
    """源图四角 20px 区域均值作为自适应图标背景色"""
    w, h = img.size
    rgb = img.convert('RGB')
    pts = []
    for cx, cy in [(10, 10), (w - 10, 10), (10, h - 10), (w - 10, h - 10)]:
        pts += [rgb.getpixel((x, y)) for x in range(max(0, cx - 10), min(w, cx + 10), 5) for y in range(max(0, cy - 10), min(h, cy + 10), 5)]
    r = sum(p[0] for p in pts) // len(pts)
    g = sum(p[1] for p in pts) // len(pts)
    b = sum(p[2] for p in pts) // len(pts)
    return '#%02X%02X%02X' % (r, g, b)

def main():
    src = sys.argv[1]
    raw = Image.open(src)
    img = square_center(trim_pure_white(raw))  # 完整内容，不裁构图
    bg = corner_bg(raw)
    print('自适应背景色:', bg)
    for dpi, (size, fg_size) in SIZES.items():
        d = os.path.join(RES, 'mipmap-' + dpi)
        os.makedirs(d, exist_ok=True)
        # legacy：完整方图原样
        img.resize((size, size), Image.LANCZOS).convert('RGBA').save(os.path.join(d, 'ic_launcher.png'), 'PNG')
        # round：完整图缩小置于圆形内（内容完整）
        canvas = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        inner = int(size * SAFE * 2 * 0.5)  # 圆直径
        content = img.resize((inner, inner), Image.LANCZOS).convert('RGBA')
        mask = Image.new('L', (inner, inner), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, inner - 1, inner - 1], fill=255)
        content.putalpha(mask)
        canvas.paste(content, ((size - inner) // 2, (size - inner) // 2), content)
        canvas.save(os.path.join(d, 'ic_launcher_round.png'), 'PNG')
        # adaptive foreground：完整图缩至安全区居中，四周透出背景色
        fcanvas = Image.new('RGBA', (fg_size, fg_size), (0, 0, 0, 0))
        fcontent = int(fg_size * SAFE)
        fcanvas.paste(img.resize((fcontent, fcontent), Image.LANCZOS).convert('RGBA'),
                      ((fg_size - fcontent) // 2, (fg_size - fcontent) // 2))
        fcanvas.save(os.path.join(d, 'ic_launcher_foreground.png'), 'PNG')
        print(' ', dpi, 'OK')
    vals = os.path.join(RES, 'values')
    with open(os.path.join(vals, 'ic_launcher_background.xml'), 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n<resources>\n    <color name="ic_launcher_background">%s</color>\n</resources>\n' % bg)
    print('ic_launcher_background.xml 已更新')

if __name__ == '__main__':
    main()
