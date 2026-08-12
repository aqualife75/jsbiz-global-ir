#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Asset preparation ops for jsbiz-global-ir (crop text bands, round corners,
whiten logos for dark covers, merge logo strips).

One op per call:
    python prep_assets.py round  in.png out.png [--radius 28]
    python prep_assets.py crop   in.png out.png --box T,B,L,R   (pixels trimmed per edge)
    python prep_assets.py whiten in.png out.png                 (dark logo -> white/transparent)
    python prep_assets.py strip  out.png --inputs a.png b.png [--gap 30] [--height 519]
    python prep_assets.py resize in.png out.png --width 800

Typical flow: crop away Korean caption bands first, then round. Verify each
result visually (Read the output image) before building the deck.
"""
import argparse

from PIL import Image, ImageDraw


def op_round(src, dst, radius):
    im = Image.open(src).convert("RGBA")
    m = Image.new("L", im.size, 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, im.width - 1, im.height - 1],
                                        radius=radius, fill=255)
    im.putalpha(m)
    im.save(dst)


def op_crop(src, dst, box):
    t, b, l, r = box
    im = Image.open(src)
    im.crop((l, t, im.width - r, im.height - b)).save(dst)


def op_whiten(src, dst):
    im = Image.open(src).convert("RGBA")
    px = im.load()
    for y in range(im.height):
        for x in range(im.width):
            r, g, b, a = px[x, y]
            lum = (r + g + b) / 3
            if lum > 200:
                px[x, y] = (255, 255, 255, 0)
            else:
                px[x, y] = (255, 255, 255, int(255 * min(1, (230 - lum) / 200)))
    im.save(dst)


def op_strip(dst, inputs, gap, height):
    ims = []
    for p in inputs:
        im = Image.open(p).convert("RGB")
        ims.append(im.resize((int(im.width * height / im.height), height)))
    w = sum(i.width for i in ims) + gap * (len(ims) - 1)
    strip = Image.new("RGB", (w, height), "white")
    x = 0
    for im in ims:
        strip.paste(im, (x, 0))
        x += im.width + gap
    strip.save(dst)


def op_resize(src, dst, width):
    im = Image.open(src)
    im.resize((width, int(im.height * width / im.width))).save(dst)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("op", choices=["round", "crop", "whiten", "strip", "resize"])
    ap.add_argument("a")
    ap.add_argument("b", nargs="?")
    ap.add_argument("--radius", type=int, default=28)
    ap.add_argument("--box", help="T,B,L,R pixels to trim")
    ap.add_argument("--inputs", nargs="+")
    ap.add_argument("--gap", type=int, default=30)
    ap.add_argument("--height", type=int, default=519)
    ap.add_argument("--width", type=int, default=800)
    args = ap.parse_args()

    if args.op == "round":
        op_round(args.a, args.b, args.radius)
    elif args.op == "crop":
        t, b, l, r = (int(v) for v in args.box.split(","))
        op_crop(args.a, args.b, (t, b, l, r))
    elif args.op == "whiten":
        op_whiten(args.a, args.b)
    elif args.op == "strip":
        op_strip(args.a, args.inputs, args.gap, args.height)
    elif args.op == "resize":
        op_resize(args.a, args.b, args.width)
    print(f"{args.op}: done")


if __name__ == "__main__":
    main()
