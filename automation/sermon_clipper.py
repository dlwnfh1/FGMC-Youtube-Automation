#!/usr/bin/env python3
"""
Download archived worship videos and export sermon-only MP4 clips.

This script expects:
- yt-dlp to be installed and available on PATH, or via --yt-dlp-bin
- ffmpeg to be installed and available on PATH, or via --ffmpeg-bin

Example:
    python automation/sermon_clipper.py --job automation/jobs.example.json
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


def slugify(value: str) -> str:
    slug = re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE).strip().lower()
    slug = re.sub(r"[-\s]+", "-", slug)
    return slug or "clip"


def parse_timecode(raw: str) -> float:
    parts = raw.strip().split(":")
    if not 1 <= len(parts) <= 3:
        raise ValueError(f"Invalid timecode: {raw}")

    try:
        numbers = [float(part) for part in parts]
    except ValueError as exc:
        raise ValueError(f"Invalid timecode: {raw}") from exc

    while len(numbers) < 3:
        numbers.insert(0, 0.0)

    hours, minutes, seconds = numbers
    return hours * 3600 + minutes * 60 + seconds


def format_timestamp(seconds: float) -> str:
    total_ms = round(seconds * 1000)
    hours, remainder = divmod(total_ms, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    secs, millis = divmod(remainder, 1000)
    return f"{hours:02}:{minutes:02}:{secs:02}.{millis:03}"


def run_command(command: list[str]) -> None:
    printable = " ".join(f'"{part}"' if " " in part else part for part in command)
    print(f"\n$ {printable}")
    subprocess.run(command, check=True)


def resolve_binary(explicit: str | None, default_name: str) -> str:
    if explicit:
        return explicit

    discovered = shutil.which(default_name)
    if not discovered:
        option_name = default_name.replace("-", "_").replace("_", "-")
        raise FileNotFoundError(
            f"Could not find '{default_name}' on PATH. "
            f"Install it or pass --{option_name}-bin."
        )
    return discovered


@dataclass
class ClipJob:
    source_url: str
    title: str
    start: float
    end: float
    speaker: str | None = None
    date: str | None = None
    browser: str | None = None
    browser_profile: str | None = None

    @property
    def slug(self) -> str:
        return slugify(self.title)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ClipJob":
        start = parse_timecode(str(payload["start"]))
        end = parse_timecode(str(payload["end"]))
        if end <= start:
            raise ValueError(
                f"Clip end must be after start for '{payload.get('title', 'untitled')}'"
            )

        return cls(
            source_url=str(payload["source_url"]).strip(),
            title=str(payload["title"]).strip(),
            start=start,
            end=end,
            speaker=str(payload["speaker"]).strip() if payload.get("speaker") else None,
            date=str(payload["date"]).strip() if payload.get("date") else None,
            browser=str(payload["browser"]).strip() if payload.get("browser") else None,
            browser_profile=(
                str(payload["browser_profile"]).strip()
                if payload.get("browser_profile")
                else None
            ),
        )


def build_cookies_args(job: ClipJob) -> list[str]:
    if not job.browser:
        return []

    browser_spec = job.browser
    if job.browser_profile:
        browser_spec = f"{browser_spec}:{job.browser_profile}"

    return ["--cookies-from-browser", browser_spec]


def ensure_downloaded(
    job: ClipJob,
    yt_dlp_bin: str,
    downloads_dir: Path,
) -> Path:
    downloads_dir.mkdir(parents=True, exist_ok=True)
    video_dir = downloads_dir / job.slug
    video_dir.mkdir(parents=True, exist_ok=True)

    existing = sorted(video_dir.glob("source.*"))
    if existing:
        print(f"Using existing download: {existing[0]}")
        return existing[0]

    output_template = str(video_dir / "source.%(ext)s")
    command = [
        yt_dlp_bin,
        "--no-progress",
        "--newline",
        "-f",
        "bv*[ext=mp4]+ba[ext=m4a]/b[ext=mp4]/b",
        "-S",
        "res:1080,vcodec:h264,acodec:m4a",
        "--merge-output-format",
        "mp4",
        *build_cookies_args(job),
        "-o",
        output_template,
        job.source_url,
    ]
    run_command(command)

    downloaded = sorted(video_dir.glob("source.*"))
    if not downloaded:
        raise FileNotFoundError(f"Download completed but no source file was found in {video_dir}")
    return downloaded[0]


def export_clip(
    job: ClipJob,
    source_file: Path,
    ffmpeg_bin: str,
    exports_dir: Path,
) -> Path:
    exports_dir.mkdir(parents=True, exist_ok=True)
    clip_name_parts = [job.title]
    if job.speaker:
        clip_name_parts.append(job.speaker)
    if job.date:
        clip_name_parts.append(job.date)
    clip_filename = f"{slugify(' '.join(clip_name_parts))}.mp4"
    destination = exports_dir / clip_filename

    command = [
        ffmpeg_bin,
        "-y",
        "-i",
        str(source_file),
        "-ss",
        format_timestamp(job.start),
        "-to",
        format_timestamp(job.end),
        "-c:v",
        "libx264",
        "-preset",
        "medium",
        "-crf",
        "18",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(destination),
    ]
    run_command(command)
    return destination


def load_jobs(job_file: Path) -> list[ClipJob]:
    payload = json.loads(job_file.read_text(encoding="utf-8"))
    if isinstance(payload, dict) and "jobs" in payload:
        raw_jobs = payload["jobs"]
    elif isinstance(payload, list):
        raw_jobs = payload
    else:
        raise ValueError("Job file must be a list or an object with a 'jobs' key")

    return [ClipJob.from_dict(item) for item in raw_jobs]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download a worship service video and export sermon clips as MP4 files."
    )
    parser.add_argument(
        "--job",
        required=True,
        help="Path to a JSON file describing one or more sermon clip jobs.",
    )
    parser.add_argument(
        "--downloads-dir",
        default="automation/downloads",
        help="Directory for downloaded full-length videos.",
    )
    parser.add_argument(
        "--exports-dir",
        default="automation/exports",
        help="Directory for final sermon MP4 clips.",
    )
    parser.add_argument(
        "--yt-dlp-bin",
        help="Explicit path to yt-dlp binary. Defaults to yt-dlp on PATH.",
    )
    parser.add_argument(
        "--ffmpeg-bin",
        help="Explicit path to ffmpeg binary. Defaults to ffmpeg on PATH.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    job_file = Path(args.job).resolve()
    downloads_dir = Path(args.downloads_dir).resolve()
    exports_dir = Path(args.exports_dir).resolve()

    try:
        yt_dlp_bin = resolve_binary(args.yt_dlp_bin, "yt-dlp")
        ffmpeg_bin = resolve_binary(args.ffmpeg_bin, "ffmpeg")
        jobs = load_jobs(job_file)
    except Exception as exc:
        print(f"Setup failed: {exc}", file=sys.stderr)
        return 1

    failures = 0
    for index, job in enumerate(jobs, start=1):
        print(f"\n[{index}/{len(jobs)}] {job.title}")
        try:
            source_file = ensure_downloaded(job, yt_dlp_bin, downloads_dir)
            destination = export_clip(job, source_file, ffmpeg_bin, exports_dir)
            print(f"Saved clip to: {destination}")
        except subprocess.CalledProcessError as exc:
            failures += 1
            print(f"Command failed with exit code {exc.returncode}", file=sys.stderr)
        except Exception as exc:
            failures += 1
            print(f"Job failed: {exc}", file=sys.stderr)

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
