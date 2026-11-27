"""Generate modern colorful icon for MyFiles (.ico and .png)."""
from pathlib import Path
from PIL import Image, ImageDraw


def create_icons(output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Rounded square vibrant indigo background
    draw.rounded_rectangle(
        [(12, 12), (size - 12, size - 12)],
        radius=52,
        fill=(67, 56, 202, 255),  # #4338CA Indigo
    )

    # Top subtle highlight arc/bar
    draw.rounded_rectangle(
        [(20, 20), (size - 20, 110)],
        radius=44,
        fill=(99, 102, 241, 120),  # Lighter indigo accent
    )

    # Folder tab (amber/gold accent)
    draw.rounded_rectangle(
        [(46, 62), (124, 104)],
        radius=16,
        fill=(245, 158, 11, 255),  # #F59E0B Amber
    )

    # Folder body (clean white)
    draw.rounded_rectangle(
        [(46, 86), (210, 202)],
        radius=22,
        fill=(255, 255, 255, 255),
    )

    # Green Excel sheet badge inside folder
    draw.rounded_rectangle(
        [(74, 108), (182, 184)],
        radius=14,
        fill=(22, 163, 74, 255),  # #16A34A Emerald/Excel Green
    )

    # Grid lines inside Excel badge
    line_color = (220, 252, 231, 255)
    # Horizontal lines
    draw.line([(90, 134), (166, 134)], fill=line_color, width=5)
    draw.line([(90, 158), (166, 158)], fill=line_color, width=5)
    # Vertical line
    draw.line([(122, 118), (122, 174)], fill=line_color, width=5)

    png_path = output_dir / "myfiles.png"
    ico_path = output_dir / "myfiles.ico"

    img.save(png_path, format="PNG")
    img.save(
        ico_path,
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    return ico_path, png_path


if __name__ == "__main__":
    assets = Path(__file__).resolve().parent / "assets"
    ico, png = create_icons(assets)
    print(f"Created icon: {ico}")
    print(f"Created png:  {png}")
