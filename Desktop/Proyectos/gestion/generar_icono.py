"""Genera icon.icns (icono de la app, estilo macOS moderno).

Uso:  python3 generar_icono.py
Requiere Pillow.  Dibuja a 2048px y reescala a 1024 para que las curvas
queden suaves; Pillow crea todos los tamaños del .icns al guardar.
"""
from PIL import Image, ImageDraw, ImageFilter

S = 2048          # lienzo de trabajo (2x del tamaño final)
FINAL = 1024

# Paleta (alineada con COLOR_HEADER_XL / acento de la app)
GRAD_TOP    = (138, 99, 175)    # #8A63AF
GRAD_BOTTOM = (78, 51, 102)     # #4E3366
ACCENT      = (108, 76, 135)    # #6C4C87
RING        = (58, 40, 80)      # #3A2850
CELL        = (233, 226, 241)   # #E9E2F1
CELL_ON     = (155, 127, 190)   # #9B7FBE
WHITE       = (255, 255, 255)


def squircle_mask(size, box, radius):
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle(box, radius=radius, fill=255)
    return mask


def sparkle(draw, cx, cy, r, color):
    """Estrella de 4 puntas (destello de limpieza)."""
    w = r * 0.24
    pts = [
        (cx, cy - r), (cx + w, cy - w), (cx + r, cy), (cx + w, cy + w),
        (cx, cy + r), (cx - w, cy + w), (cx - r, cy), (cx - w, cy - w),
    ]
    draw.polygon(pts, fill=color)


def main():
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # --- Squircle de fondo con degradado vertical ---
    margin = int(S * 0.098)
    box = (margin, margin, S - margin, S - margin)
    radius = int(S * 0.225)

    grad = Image.new("RGBA", (S, S))
    gdraw = ImageDraw.Draw(grad)
    for y in range(S):
        t = y / S
        c = tuple(int(a + (b - a) * t) for a, b in zip(GRAD_TOP, GRAD_BOTTOM))
        gdraw.line([(0, y), (S, y)], fill=c + (255,))
    img.paste(grad, (0, 0), squircle_mask(S, box, radius))

    # --- Sombra de la tarjeta ---
    card = (int(S * 0.20), int(S * 0.26), int(S * 0.80), int(S * 0.82))
    card_r = int(S * 0.06)
    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(
        [card[0], card[1] + int(S * 0.012)] + [card[2], card[3] + int(S * 0.012)],
        radius=card_r, fill=(20, 8, 34, 90),
    )
    img = Image.alpha_composite(img, shadow.filter(ImageFilter.GaussianBlur(S * 0.014)))

    draw = ImageDraw.Draw(img)

    # --- Tarjeta de calendario ---
    draw.rounded_rectangle(card, radius=card_r, fill=WHITE + (255,))

    # Cabecera de la tarjeta (esquinas superiores redondeadas)
    header_h = int(S * 0.14)
    draw.rounded_rectangle(
        [card[0], card[1], card[2], card[1] + header_h + card_r],
        radius=card_r, fill=ACCENT + (255,),
    )
    draw.rectangle(
        [card[0], card[1] + header_h, card[2], card[1] + header_h + card_r],
        fill=WHITE + (255,),
    )
    draw.rectangle(
        [card[0], card[1] + int(header_h * 0.55), card[2], card[1] + header_h],
        fill=ACCENT + (255,),
    )

    # --- Anillas ---
    ring_w = int(S * 0.030)
    ring_h = int(S * 0.115)
    ring_y = card[1] - int(ring_h * 0.55)
    for cx in (int(S * 0.355), int(S * 0.645)):
        draw.rounded_rectangle(
            [cx - ring_w, ring_y, cx + ring_w, ring_y + ring_h],
            radius=ring_w, fill=RING + (255,),
        )

    # --- Retícula de días (4 x 3), con una diagonal encendida = rotación ---
    cols, rows = 4, 3
    grid_x0, grid_x1 = card[0] + int(S * 0.055), card[2] - int(S * 0.055)
    grid_y0, grid_y1 = card[1] + header_h + int(S * 0.055), card[3] - int(S * 0.055)
    cell_gap = int(S * 0.022)
    cw = (grid_x1 - grid_x0 - cell_gap * (cols - 1)) // cols
    ch = (grid_y1 - grid_y0 - cell_gap * (rows - 1)) // rows
    encendidas = {(0, 0), (1, 1), (2, 2)}
    for r in range(rows):
        for c in range(cols):
            x = grid_x0 + c * (cw + cell_gap)
            y = grid_y0 + r * (ch + cell_gap)
            color = CELL_ON if (r, c) in encendidas else CELL
            draw.rounded_rectangle(
                [x, y, x + cw, y + ch], radius=int(S * 0.018), fill=color + (255,)
            )

    # --- Insignia con destello (limpieza) ---
    bx, by, br = int(S * 0.76), int(S * 0.78), int(S * 0.115)
    draw.ellipse([bx - br, by - br, bx + br, by + br], fill=ACCENT + (255,),
                 outline=WHITE + (255,), width=int(S * 0.012))
    sparkle(draw, bx - int(br * 0.10), by, int(br * 0.52), WHITE + (255,))
    sparkle(draw, bx + int(br * 0.45), by - int(br * 0.45), int(br * 0.22), WHITE + (255,))

    # --- Brillo superior sutil ---
    gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(gloss).ellipse(
        [-int(S * 0.25), -int(S * 0.55), int(S * 1.25), int(S * 0.45)],
        fill=(255, 255, 255, 26),
    )
    gloss = gloss.filter(ImageFilter.GaussianBlur(S * 0.04))
    masked_gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    masked_gloss.paste(gloss, (0, 0), squircle_mask(S, box, radius))
    img = Image.alpha_composite(img, masked_gloss)

    final = img.resize((FINAL, FINAL), Image.LANCZOS)
    final.save("icon.icns")
    final.save("icon_preview.png")
    print("icon.icns regenerado")


if __name__ == "__main__":
    main()
