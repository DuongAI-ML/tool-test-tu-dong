"""Vẽ icon app (ô vuông xanh bo góc + dấu tích trắng, giống web/icon.svg) → web/app.ico. Chỉ chạy lúc đóng gói."""
import os

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def draw(size):
    s = 8                                   # vẽ to rồi thu nhỏ cho mịn
    big = Image.new("RGBA", (64 * s, 64 * s), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    d.rounded_rectangle((0, 0, 64 * s - 1, 64 * s - 1), radius=14 * s, fill=(29, 78, 216, 255))
    pts = [(18 * s, 33 * s), (27 * s, 42 * s), (46 * s, 22 * s)]
    d.line(pts, fill=(255, 255, 255, 255), width=7 * s, joint="curve")
    for x, y in (pts[0], pts[-1]):
        r = 3.5 * s
        d.ellipse((x - r, y - r, x + r, y + r), fill=(255, 255, 255, 255))
    return big.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    out = os.path.join(ROOT, "web", "app.ico")
    draw(256).save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("Đã tạo", out)
