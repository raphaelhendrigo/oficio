import argparse
import html
import re
from pathlib import Path


def extract_timestamp(name: str) -> str:
    m = re.search(r"t=([0-9]{2}-[0-9]{2}-[0-9]{2}\.[0-9]{3})", name)
    return m.group(1) if m else ""


def timestamp_to_seconds(ts: str) -> float:
    if not ts:
        return 0.0
    parts = ts.split("-")
    if len(parts) != 3:
        return 0.0
    h = int(parts[0])
    m = int(parts[1])
    s_part = parts[2]
    if "." in s_part:
        s, ms = s_part.split(".", 1)
        return h * 3600 + m * 60 + int(s) + (int(ms) / 1000.0)
    return h * 3600 + m * 60 + int(s_part)


def gather_frames(frames_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*.png" if recursive else "*.png"
    return [p for p in frames_dir.glob(pattern) if p.is_file()]


def build_html(frames: list[Path], frames_dir: Path) -> str:
    items = []
    for p in frames:
        rel = p.relative_to(frames_dir)
        ts = extract_timestamp(p.name)
        items.append(
            {
                "path": rel.as_posix(),
                "name": p.name,
                "timestamp": ts,
                "seconds": timestamp_to_seconds(ts),
            }
        )

    items.sort(key=lambda x: (x["seconds"], x["name"]))

    parts = [
        "<!doctype html>",
        "<html lang='en'>",
        "<head>",
        "  <meta charset='utf-8'/>",
        "  <meta name='viewport' content='width=device-width, initial-scale=1'/>",
        "  <title>Video Frames Contact Sheet</title>",
        "  <style>",
        "    body { font-family: Arial, sans-serif; margin: 16px; background: #111; color: #eee; }",
        "    .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 12px; }",
        "    .card { background: #1c1c1c; border: 1px solid #333; border-radius: 6px; padding: 8px; }",
        "    img { width: 100%; height: auto; display: block; border-radius: 4px; }",
        "    .meta { font-size: 12px; margin-top: 6px; color: #bbb; }",
        "  </style>",
        "</head>",
        "<body>",
        f"<h1>Video Frames ({len(items)})</h1>",
        "<div class='grid'>",
    ]

    for item in items:
        name = html.escape(item["name"])
        path = html.escape(item["path"])
        ts = html.escape(item["timestamp"] or "n/a")
        parts.extend(
            [
                "<div class='card'>",
                f"  <img src='{path}' loading='lazy' />",
                f"  <div class='meta'>{ts} - {name}</div>",
                "</div>",
            ]
        )

    parts.extend(["</div>", "</body>", "</html>"])
    return "\n".join(parts)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate an HTML contact sheet for extracted frames.")
    parser.add_argument(
        "--frames-dir",
        default="docs/video_frames",
        help="Directory containing extracted frames.",
    )
    parser.add_argument(
        "--output",
        default="docs/video_frames/index.html",
        help="Output HTML file.",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Do not scan subfolders.",
    )
    args = parser.parse_args()

    frames_dir = Path(args.frames_dir)
    if not frames_dir.exists():
        print(f"ERROR: frames directory not found: {frames_dir}")
        return

    frames = gather_frames(frames_dir, recursive=not args.no_recursive)
    if not frames:
        print("No frames found. Run tools/extract_frames.py first.")
        return

    html_text = build_html(frames, frames_dir)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_text, encoding="utf-8")
    print(f"Contact sheet written: {output_path}")


if __name__ == "__main__":
    main()
