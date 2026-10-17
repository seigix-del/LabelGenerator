"""Draws the app icon (stacked cartons with an IN/OUT arrow) and writes build/icon.png + build/icon.icns.

    python desktop/build/make_icon.py
"""
import pathlib
from PIL import Image, ImageDraw

OUT = pathlib.Path(__file__).resolve().parent
S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# macOS icon grid: 824px rounded square centred on a 1024 canvas
pad, r = 100, 185
d.rounded_rectangle([pad, pad, S - pad, S - pad], r, fill=(11, 110, 127))
d.rounded_rectangle([pad, pad + 412, S - pad, S - pad], r, fill=(9, 92, 106))
d.rectangle([pad, pad + 412, S - pad, pad + 412 + r], fill=(9, 92, 106))

# three cartons
cream, tape, edge = (246, 234, 210), (214, 186, 132), (190, 160, 105)
for x, y, w, h in [(250, 520, 250, 210), (524, 520, 250, 210), (387, 300, 250, 210)]:
    d.rounded_rectangle([x, y, x + w, y + h], 18, fill=cream, outline=edge, width=8)
    d.rectangle([x + w // 2 - 18, y, x + w // 2 + 18, y + h], fill=tape)

# green IN arrow (left, down-right) and red OUT arrow (right, up)
def arrow(points, color):
    d.line(points[:2], fill=color, width=46)
    d.polygon(points[2:], fill=color)

arrow([(200, 300), (300, 410), (330, 450), (250, 420), (330, 360)], (72, 194, 124))
arrow([(824, 410), (824, 290), (824, 230), (774, 300), (874, 300)], (240, 112, 93))

img.save(OUT / "icon.png")
img.save(OUT / "icon.icns", sizes=[(16, 16), (32, 32), (64, 64), (128, 128), (256, 256), (512, 512), (1024, 1024)])
print("icon written")
