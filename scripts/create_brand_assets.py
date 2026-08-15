"""Generate deterministic PDF2MD Windows icon assets."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def load_font(size: int, bold: bool = False):
    font_name = "JetBrainsMono-Bold.ttf" if bold else "Outfit-Variable.ttf"
    font_path = ASSETS / "fonts" / font_name
    if font_path.is_file():
        return ImageFont.truetype(str(font_path), size)
    return ImageFont.load_default()


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    size = 256
    image = Image.new("RGBA", (size, size), (10, 10, 15, 255))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((18, 18, 238, 238), radius=48, fill=(18, 18, 26), outline=(255, 107, 53), width=8)
    draw.rounded_rectangle((65, 42, 191, 214), radius=12, fill=(232, 232, 240))
    draw.polygon([(155, 42), (191, 78), (155, 78)], fill=(255, 143, 94))
    for y, width in ((100, 82), (122, 82), (144, 62)):
        draw.rounded_rectangle((87, y, 87 + width, y + 7), radius=3, fill=(42, 42, 58))
    draw.rounded_rectangle((48, 164, 208, 224), radius=14, fill=(255, 107, 53))
    label = "MD"
    font = load_font(43, bold=True)
    bounds = draw.textbbox((0, 0), label, font=font)
    x = (size - (bounds[2] - bounds[0])) / 2
    draw.text((x, 168), label, font=font, fill=(255, 255, 255))
    image.save(ASSETS / "pdf2md.png")
    image.save(
        ASSETS / "pdf2md.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )


if __name__ == "__main__":
    main()
