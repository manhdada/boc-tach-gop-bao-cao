from pathlib import Path
import sys

from PIL import Image, ImageDraw


def main(folder: str, batch_size: int = 4):
    root = Path(folder)
    pages = sorted(root.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
    for batch_index in range(0, len(pages), batch_size):
        batch = pages[batch_index:batch_index + batch_size]
        opened = [Image.open(path).convert("RGB") for path in batch]
        scale = 0.5
        thumbs = [image.resize((int(image.width * scale), int(image.height * scale))) for image in opened]
        width = max(image.width for image in thumbs) + 40
        height = sum(image.height + 55 for image in thumbs) + 20
        sheet = Image.new("RGB", (width, height), "#d8d8d8")
        draw = ImageDraw.Draw(sheet)
        y = 15
        for path, image in zip(batch, thumbs):
            draw.text((20, y), path.stem, fill="black")
            y += 25
            sheet.paste(image, (20, y))
            y += image.height + 30
        first = batch_index + 1
        last = batch_index + len(batch)
        sheet.save(root / f"sheet-{first:02d}-{last:02d}.png")


if __name__ == "__main__":
    main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 4)
