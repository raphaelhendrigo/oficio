import argparse
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path


def format_timestamp_hms(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    total = int(seconds)
    hours = total // 3600
    minutes = (total % 3600) // 60
    secs = total % 60
    return f"{hours:02d}-{minutes:02d}-{secs:02d}"


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


def unique_name(base: str, used: dict[str, int]) -> str:
    if base not in used:
        used[base] = 1
        return base
    used[base] += 1
    return f"{base}_{used[base]:02d}"


def clean_pngs(out_dir: Path) -> None:
    if not out_dir.exists():
        return
    for p in out_dir.glob("*.png"):
        try:
            p.unlink()
        except Exception:
            pass


def extract_frames_ffmpeg(
    video_path: Path,
    out_dir: Path,
    every: float,
    manifest_path: Path | None,
) -> int:
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        raise RuntimeError("ffmpeg not found")

    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_dir = out_dir / f".tmp_extract_{int(time.time())}"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    out_pattern = tmp_dir / "frame_%06d.png"

    fps_expr = f"fps=1/{every}" if every != 1 else "fps=1"
    vf = f"{fps_expr},showinfo"
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
    proc = subprocess.run(args, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "ffmpeg failed")

    pts_times = parse_pts_times(proc.stderr)
    files = sorted(tmp_dir.glob("frame_*.png"))
    if not files:
        return 0

    if len(pts_times) != len(files):
        pts_times = [(idx * every) for idx in range(len(files))]

    used: dict[str, int] = {}
    manifest = []
    for src, pts_time in zip(files, pts_times):
        ts = format_timestamp_hms(pts_time)
        base = unique_name(ts, used)
        dest = out_dir / f"{base}.png"
        if dest.exists():
            dest.unlink()
        try:
            src.rename(dest)
        except Exception:
            dest.write_bytes(src.read_bytes())
            src.unlink(missing_ok=True)
        manifest.append({"file": dest.name, "timestamp": ts, "pts_time": pts_time})

    for p in files:
        try:
            p.unlink(missing_ok=True)
        except Exception:
            pass
    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    if manifest_path:
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    return len(manifest)


def extract_frames_opencv(
    video_path: Path,
    out_dir: Path,
    every: float,
    manifest_path: Path | None,
) -> int:
    try:
        import cv2  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"opencv-python not available: {exc}") from exc

    out_dir.mkdir(parents=True, exist_ok=True)
    cap = cv2.VideoCapture(str(video_path))
    if not cap.isOpened():
        raise RuntimeError("Failed to open video with OpenCV")

    fps = cap.get(cv2.CAP_PROP_FPS) or 0.0
    frame_count = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0.0
    duration = frame_count / fps if fps > 0 else 0.0
    if duration <= 0:
        duration = 0.0

    used: dict[str, int] = {}
    manifest = []
    t = 0.0
    idx = 0
    while True:
        if duration and t > duration + 0.5:
            break
        cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000.0)
        ok, frame = cap.read()
        if not ok:
            break
        ts = format_timestamp_hms(t)
        base = unique_name(ts, used)
        dest = out_dir / f"{base}.png"
        if not cv2.imwrite(str(dest), frame):
            break
        manifest.append({"file": dest.name, "timestamp": ts, "pts_time": t})
        idx += 1
        t += every

    cap.release()
    if manifest_path:
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return idx


def reuse_existing_frames(
    source_dir: Path,
    out_dir: Path,
    manifest_path: Path | None,
) -> int:
    manifest_file = source_dir / "interval_frames_manifest.json"
    items = []
    if manifest_file.exists():
        try:
            items = json.loads(manifest_file.read_text(encoding="utf-8"))
        except Exception:
            items = []
    if not items:
        for p in sorted(source_dir.glob("interval_frame_*_t=*.png")):
            m = re.search(r"t=([0-9]{2}-[0-9]{2}-[0-9]{2})", p.name)
            if not m:
                continue
            items.append({"file": p.name, "timestamp": m.group(1), "pts_time": None})

    if not items:
        raise RuntimeError("No existing frames found to reuse")

    out_dir.mkdir(parents=True, exist_ok=True)
    used: dict[str, int] = {}
    manifest = []
    skipped = 0
    for item in items:
        src = source_dir / item["file"]
        ts_full = item.get("timestamp", "")
        ts = ts_full.split(".", 1)[0] if ts_full else ""
        if not ts:
            continue
        base = unique_name(ts, used)
        dest = out_dir / f"{base}.png"
        if dest.exists():
            manifest.append({"file": dest.name, "timestamp": ts, "pts_time": item.get("pts_time")})
            continue
        try:
            shutil.copy2(src, dest)
        except PermissionError:
            try:
                dest.write_bytes(src.read_bytes())
            except Exception:
                skipped += 1
                continue
        except Exception:
            skipped += 1
            continue
        manifest.append({"file": dest.name, "timestamp": ts, "pts_time": item.get("pts_time")})

    if manifest_path:
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    if skipped:
        print(f"Warning: skipped {skipped} frame(s) due to copy errors")
    return len(manifest)


def write_metadata(video_path: Path, metadata_path: Path) -> None:
    size_bytes = video_path.stat().st_size
    video_abs = video_path.resolve()
    data = {
        "path": str(video_path),
        "size_bytes": size_bytes,
        "duration_seconds": None,
        "duration_hms": None,
        "width": None,
        "height": None,
        "frame_rate_fps": None,
        "bit_rate_bps": None,
        "source": "windows_shell",
    }
    try:
        import win32com.client  # type: ignore

        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.Namespace(str(video_abs.parent))
        item = folder.ParseName(video_abs.name) if folder else None
        if not item:
            raise RuntimeError("Shell namespace unavailable")

        def prop(name: str):
            try:
                return item.ExtendedProperty(name)
            except Exception:
                return None

        width = prop("System.Video.FrameWidth")
        height = prop("System.Video.FrameHeight")
        fps_raw = prop("System.Video.FrameRate")
        duration_raw = prop("System.Media.Duration")
        bitrate = prop("System.Video.TotalBitrate")

        if width:
            data["width"] = int(width)
        if height:
            data["height"] = int(height)
        if fps_raw:
            fps_val = float(fps_raw)
            data["frame_rate_fps"] = round(fps_val / 1000.0, 3) if fps_val > 1000 else fps_val
        if duration_raw:
            duration_sec = float(duration_raw) / 10_000_000.0
            data["duration_seconds"] = round(duration_sec, 3)
            h = int(duration_sec // 3600)
            m = int((duration_sec % 3600) // 60)
            s = int(duration_sec % 60)
            data["duration_hms"] = f"{h:02d}:{m:02d}:{s:02d}"
        if bitrate:
            data["bit_rate_bps"] = int(bitrate)
    except Exception:
        data["source"] = "fallback_file_only"

    metadata_path.parent.mkdir(parents=True, exist_ok=True)
    metadata_path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Ingest a local video into frames and metadata.")
    parser.add_argument("--video", default="video_silvana.mp4", help="Path to the input video file.")
    parser.add_argument("--out-dir", default="docs/video_silvana/frames", help="Output directory for frames.")
    parser.add_argument("--every", type=float, default=1.0, help="Interval in seconds for frame capture.")
    parser.add_argument(
        "--reuse-existing",
        action="store_true",
        help="Reuse already extracted frames from docs/video_frames if available.",
    )
    parser.add_argument(
        "--existing-frames-dir",
        default="docs/video_frames",
        help="Directory containing previously extracted frames.",
    )
    parser.add_argument(
        "--manifest",
        default="docs/video_silvana/frames_manifest.json",
        help="Output manifest JSON path.",
    )
    parser.add_argument(
        "--metadata",
        default="docs/video_silvana/metadata.json",
        help="Output metadata JSON path.",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Do not delete existing frames in the output directory.",
    )
    args = parser.parse_args()

    video_path = Path(args.video)
    if not video_path.exists():
        print(f"ERROR: video not found: {video_path}")
        sys.exit(2)

    out_dir = Path(args.out_dir)
    if not args.no_clean:
        clean_pngs(out_dir)

    manifest_path = Path(args.manifest) if args.manifest else None
    if args.reuse_existing:
        count = reuse_existing_frames(Path(args.existing_frames_dir), out_dir, manifest_path)
        write_metadata(video_path, Path(args.metadata))
        print(f"Reused {count} frames into {out_dir}")
        return

    try:
        count = extract_frames_ffmpeg(video_path, out_dir, args.every, manifest_path)
    except Exception:
        count = None

    if count is None:
        try:
            count = extract_frames_opencv(video_path, out_dir, args.every, manifest_path)
        except Exception as exc:
            print(f"ERROR: {exc}")
            sys.exit(2)

    write_metadata(video_path, Path(args.metadata))
    print(f"Extracted {count} frames into {out_dir}")


if __name__ == "__main__":
    main()
