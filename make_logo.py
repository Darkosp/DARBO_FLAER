# -*- coding: utf-8 -*-
"""Го генерира лого-написот DARBO GROUP како PNG.

Michroma има само една тежина (regular), а логото од понудата е подебело
(BankGothic Md BT Medium). Затоа буквите се цртаат со контура од истата боја —
измерено: референцата има однос дебелина/висина 0,200, Michroma regular 0,137.

Резултатот е `assets/logo-wordmark.png` (транспарентна позадина, ~1500 dpi при
печатење во големина од 7,3 cm).
"""
from PIL import Image, ImageDraw, ImageFont

TTF   = "assets/fonts/Michroma-Regular.ttf"
OUT   = "assets/logo-wordmark.png"
SIZE  = 400          # px — голем, за да нема загуба при печатење
STROKE = 8           # px — ~0,19 дебелина/висина (логото од понудата: ~0,20)
GAP_RATIO = 0.50     # јаз меѓу DARBO и GROUP, како дел од висината на капите

GREEN = (26, 122, 71)    # --color-darbo  #1A7A47
RED   = (210, 58, 42)    # --color-group  #D23A2A


def main():
    font = ImageFont.truetype(TTF, SIZE)
    bbox = font.getbbox("D")
    cap = bbox[3] - bbox[1]

    measure = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    w_darbo = measure.textlength("DARBO", font=font)
    w_group = measure.textlength("GROUP", font=font)
    gap = cap * GAP_RATIO
    pad = STROKE * 2 + 8

    width = int(w_darbo + gap + w_group) + pad * 2
    height = cap + pad * 2
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # PIL го позиционира текстов по неговиот ascent; порамнуваме по капите
    y_text = pad - bbox[1]
    x = pad
    for text, color, w in (("DARBO", GREEN, w_darbo), ("GROUP", RED, w_group)):
        draw.text((x, y_text), text, font=font, fill=color,
                  stroke_width=STROKE, stroke_fill=color)
        x += w + gap

    bbox_img = img.getbbox()
    img = img.crop(bbox_img)
    img.save(OUT)
    print(f"saved {OUT}: {img.size[0]}x{img.size[1]} px, cap {cap} px, "
          f"aspect {img.size[0]/img.size[1]:.4f}")


if __name__ == "__main__":
    main()
