import argparse
import json
import os
import re
import sys
from pathlib import Path


def extract_timestamp(name: str) -> str:
    m = re.search(r"t=([0-9]{2}-[0-9]{2}-[0-9]{2}\.[0-9]{3})", name)
    return m.group(1) if m else ""


def main() -> None:
    parser = argparse.ArgumentParser(description="Run OCR on extracted frames (optional).")
    parser.add_argument(
        "--frames-dir",
        default="docs/video_frames",
        help="Directory containing extracted frames.",
    )
    parser.add_argument(
        "--output",
        default="docs/video_frames/ocr.json",
        help="Output JSON file.",
    )
    parser.add_argument(
        "--lang",
        default="por",
        help="Tesseract language (e.g., por, eng, por+eng).",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Limit number of frames (0 = no limit).",
    )
    args = parser.parse_args()

    try:
        from PIL import Image  # type: ignore
        import pytesseract  # type: ignore
    except Exception:
        print("OCR dependencies not available.")
        print("Install with: pip install pytesseract pillow")
        print("Also install Tesseract OCR and ensure it is in PATH.")
        print("Windows installer: https://github.com/UB-Mannheim/tesseract/wiki")
        sys.exit(2)

    tesseract_cmd = os.getenv("TESSERACT_CMD")
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    frames_dir = Path(args.frames_dir)
    if not frames_dir.exists():
        print(f"ERROR: frames directory not found: {frames_dir}")
        return

    frames = sorted(frames_dir.glob("**/*.png"))
    if args.limit and args.limit > 0:
        frames = frames[: args.limit]
    if not frames:
        print("No frames found. Run tools/extract_frames.py first.")
        return

    results = []
    for idx, frame in enumerate(frames, start=1):
        try:
            img = Image.open(frame)
            text = pytesseract.image_to_string(img, lang=args.lang)
        except Exception as exc:
            text = ""
            print(f"Warning: OCR failed for {frame.name}: {exc}")
        results.append(
            {
                "file": frame.name,
                "timestamp": extract_timestamp(frame.name),
                "text": text,
            }
        )
        if idx % 25 == 0:
            print(f"OCR progress: {idx}/{len(frames)}")

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
    print(f"OCR output written: {output_path}")


if __name__ == "__main__":
    main()
