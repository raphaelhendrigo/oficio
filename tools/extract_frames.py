import argparse
import json
import re
import shutil
import subprocess
import sys
import time
from datetime import timedelta
from pathlib import Path


def format_timestamp(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    td = timedelta(seconds=seconds)
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    secs = total_seconds % 60
    millis = int(round((seconds - int(seconds)) * 1000.0))
    return f"{hours:02d}-{minutes:02d}-{secs:02d}.{millis:03d}"


def parse_pts_times(ffmpeg_stderr: str) -> list[float]:
    pts = []
    for line in ffmpeg_stderr.splitlines():
        m = re.search(r"pts_time:([0-9]+(?:\.[0-9]+)?)", line)
        if m:
            try:
                pts.append(float(m.group(1)))
            except ValueError:
                continue
    return pts


def build_vf(mode: str, every: float, scene_threshold: float) -> str:
    if mode == "interval":
        fps_expr = f"fps=1/{every}" if every != 1 else "fps=1"
        return f"{fps_expr},showinfo"
    if mode == "scene":
        return f"select='gt(scene,{scene_threshold})',showinfo"
    raise ValueError(f"Unknown mode: {mode}")


def clean_existing(out_dir: Path, prefix: str) -> None:
    for p in out_dir.glob(f"{prefix}frame_*_t=*.png"):
        try:
            p.unlink()
        except Exception:
            pass


def extract_frames(
    video_path: Path,
    out_dir: Path,
    mode: str,
    every: float,
    scene_threshold: float,
    prefix: str,
    clean: bool,
    write_manifest: bool,
) -> None:
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        print("ERROR: ffmpeg not found. Install ffmpeg and ensure it is in PATH.")
        print("Tip: https://ffmpeg.org/download.html (Windows builds are available).")
        sys.exit(2)

    out_dir.mkdir(parents=True, exist_ok=True)
    if clean:
        clean_existing(out_dir, prefix)

    tmp_dir = out_dir / f".tmp_extract_{prefix}{int(time.time())}"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    out_pattern = tmp_dir / f"{prefix}frame_%06d.png"

    vf = build_vf(mode, every, scene_threshold)
    args = [
        ffmpeg,
        "-hide_banner",
        "-i",
        str(video_path),
        "-vf",
        vf,
        "-vsync",
        "vfr",
        str(out_pattern),
    ]
    print(f"Running: {' '.join(args)}")
    proc = subprocess.run(args, capture_output=True, text=True)
    if proc.returncode != 0:
        print(proc.stderr)
        print("ERROR: ffmpeg failed.")
        sys.exit(proc.returncode)

    pts_times = parse_pts_times(proc.stderr)
    files = sorted(tmp_dir.glob(f"{prefix}frame_*.png"))
    if not files:
        print("No frames extracted. Check the video path and parameters.")
        return

    if len(pts_times) != len(files):
        print(
            f"Warning: pts_time count ({len(pts_times)}) != frame count ({len(files)}). "
            "Falling back to index-based timestamps."
        )
        pts_times = [(idx * every) for idx in range(len(files))]

    manifest = []
    for idx, (src, pts_time) in enumerate(zip(files, pts_times), start=1):
        ts = format_timestamp(pts_time)
        dest_name = f"{prefix}frame_{idx:06d}_t={ts}.png"
        dest = out_dir / dest_name
        try:
            if dest.exists():
                dest.unlink()
            src.rename(dest)
        except Exception:
            dest = out_dir / dest_name
            dest.write_bytes(src.read_bytes())
            src.unlink(missing_ok=True)
        manifest.append(
            {
                "file": dest.name,
                "timestamp": ts,
                "pts_time": pts_time,
                "mode": mode,
            }
        )

    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    if write_manifest:
        manifest_path = out_dir / f"{prefix}frames_manifest.json"
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        print(f"Manifest written: {manifest_path}")

    print(f"Extracted {len(manifest)} frames to {out_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract frames from a video (interval or scene detection).")
    parser.add_argument(
        "--video",
        default="video_silvana.mp4",
        help="Path to the input video file.",
    )
    parser.add_argument(
        "--out-dir",
        default="docs/video_frames",
        help="Output directory for frames.",
    )
    parser.add_argument(
        "--mode",
        choices=["interval", "scene", "both"],
        default="interval",
        help="Extraction mode.",
    )
    parser.add_argument(
        "--every",
        type=float,
        default=1.0,
        help="Interval in seconds for interval mode.",
    )
    parser.add_argument(
        "--scene-threshold",
        type=float,
        default=0.30,
        help="Scene detection threshold (higher = fewer frames).",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Do not delete existing frames with the same prefix.",
    )
    parser.add_argument(
        "--no-manifest",
        action="store_true",
        help="Do not write frames manifest JSON.",
    )
    args = parser.parse_args()

    video_path = Path(args.video)
    if not video_path.exists():
        print(f"ERROR: video not found: {video_path}")
        sys.exit(2)

    out_dir = Path(args.out_dir)
    clean = not args.no_clean
    write_manifest = not args.no_manifest

    if args.mode in ("interval", "both"):
        extract_frames(
            video_path,
            out_dir,
            mode="interval",
            every=args.every,
            scene_threshold=args.scene_threshold,
            prefix="interval_",
            clean=clean,
            write_manifest=write_manifest,
        )
    if args.mode in ("scene", "both"):
        extract_frames(
            video_path,
            out_dir,
            mode="scene",
            every=args.every,
            scene_threshold=args.scene_threshold,
            prefix="scene_",
            clean=clean,
            write_manifest=write_manifest,
        )


if __name__ == "__main__":
    main()
