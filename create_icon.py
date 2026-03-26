"""
OutlookGemini 아이콘 생성 스크립트
파란 배경 + 흰색 편지봉투 + 별(AI) 모양
"""
from PIL import Image, ImageDraw, ImageFont
import math


def draw_rounded_rect(draw, xy, radius, fill):
    x0, y0, x1, y1 = xy
    draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
    draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
    draw.ellipse([x0, y0, x0 + radius * 2, y0 + radius * 2], fill=fill)
    draw.ellipse([x1 - radius * 2, y0, x1, y0 + radius * 2], fill=fill)
    draw.ellipse([x0, y1 - radius * 2, x0 + radius * 2, y1], fill=fill)
    draw.ellipse([x1 - radius * 2, y1 - radius * 2, x1, y1], fill=fill)


def draw_envelope(draw, cx, cy, w, h, color, lw):
    """편지봉투 외곽선"""
    x0, y0 = cx - w // 2, cy - h // 2
    x1, y1 = cx + w // 2, cy + h // 2
    # 봉투 박스
    draw.rectangle([x0, y0, x1, y1], outline=color, width=lw)
    # V자 접힘선 (위)
    draw.line([x0, y0, cx, cy - h // 8], fill=color, width=lw)
    draw.line([x1, y0, cx, cy - h // 8], fill=color, width=lw)


def draw_sparkle(draw, cx, cy, r, color, lw):
    """4방향 별 모양"""
    for angle in range(0, 360, 45):
        rad = math.radians(angle)
        length = r if angle % 90 == 0 else r * 0.45
        x = cx + math.cos(rad) * length
        y = cy + math.sin(rad) * length
        draw.line([cx, cy, x, y], fill=color, width=lw)


def make_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    pad = size * 0.06
    radius = size * 0.22
    bg_color = (26, 115, 232, 255)       # #1a73e8
    fg_color = (255, 255, 255, 255)      # white
    accent  = (255, 213, 79, 255)        # 노란 별

    # 배경 둥근 사각형
    draw_rounded_rect(draw, [pad, pad, size - pad, size - pad],
                      int(radius), bg_color)

    # 편지봉투
    lw = max(2, size // 32)
    ew = int(size * 0.52)
    eh = int(size * 0.36)
    draw_envelope(draw, size // 2, int(size * 0.54), ew, eh, fg_color, lw)

    # 오른쪽 위 반짝이
    sr = int(size * 0.13)
    slw = max(2, size // 40)
    draw_sparkle(draw, int(size * 0.72), int(size * 0.28), sr, accent, slw)

    return img


# 여러 크기 생성 후 .ico 저장
sizes = [16, 24, 32, 48, 64, 128, 256]
images = [make_icon(s) for s in sizes]
images[0].save(
    "icon.ico",
    format="ICO",
    sizes=[(s, s) for s in sizes],
    append_images=images[1:],
)
print("icon.ico 생성 완료")
