#!/usr/bin/env python3
"""
Windows desktop app for downloading archived worship videos and exporting sermon clips.
"""

from __future__ import annotations

import base64
import audioop
import hashlib
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import wave
import winreg
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent


def resolve_project_dir() -> Path:
    if not getattr(sys, "frozen", False):
        return APP_DIR.parent

    # For the installed/portable EXE, keep all data next to the EXE so the package
    # can be copied to another computer without accidentally using a parent folder.
    return APP_DIR


PROJECT_DIR = resolve_project_dir()
RUNTIME_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else PROJECT_DIR
WORK_DIR = PROJECT_DIR / "desktop_app_data"
DOWNLOADS_DIR = WORK_DIR / "downloads"
EXPORTS_DIR = WORK_DIR / "exports"
PREVIEW_DIR = WORK_DIR / "preview_audio"
TRANSCRIPT_DIR = WORK_DIR / "transcripts"
FRAME_DIR = WORK_DIR / "frames"
REVIEW_DIR = WORK_DIR / "review_clips"
SETTINGS_PATH = WORK_DIR / "settings.json"
YOUTUBE_TOKEN_PATH = WORK_DIR / "youtube_token.json"

DEFAULT_MARKERS = [
    "\ub2e4\ud568\uaed8 \uae30\ub3c4 \ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
    "\ub2e4\ud568\uaed8 \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\uace0\uac1c \uc219\uc5ec \uae30\ub3c4 \ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\uace0\uac1c\uc219\uc5ec \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4 \ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4\ud558\uc2dc\uaca0\uc2b5\ub2c8\ub2e4",
    "\uc131\uac00\ub300 \ucc2c\uc591",
    "\ud2b9\uc1a1",
    "\ubc15\uc218",
    "\ud558\ub098\ub2d8\uc758 \ub9d0\uc500",
    "\uc624\ub298 \ub9d0\uc500",
    "\ubcf8\ubb38 \ub9d0\uc500",
    "\uc131\uacbd \ub9d0\uc500",
    "\ub9d0\uc500\uc744 \ubcf4\uaca0\uc2b5\ub2c8\ub2e4",
    "\ud568\uaed8 \ub9d0\uc500",
    "\uc124\uad50 \ub9d0\uc500",
    "\uc624\ub298 \ubcf8\ubb38",
    "\uc624\ub298 \ub098\ub20c \ub9d0\uc500",
]

DEFAULT_END_MARKERS = [
    "\ucd95\uc6d0 \ub4dc\ub9bd\ub2c8\ub2e4",
    "\ucd95\uc6d0\ud569\ub2c8\ub2e4",
    "\ucd95\uac74\ud569\ub2c8\ub2e4",
    "\ucd95\uac74 \ub4dc\ub9bd\ub2c8\ub2e4",
    "\ucd94\uac74\ud569\ub2c8\ub2e4",
    "\ucd94\uac74 \ub4dc\ub9bd\ub2c8\ub2e4",
    "\ucd94\uad8c\ud569\ub2c8\ub2e4",
    "\ucd94\uad6c\ud569\ub2c8\ub2e4",
]

DEFAULT_TRANSCRIPTION_SKIP_START_SECONDS = 10 * 60
DEFAULT_TRANSCRIPTION_SKIP_END_SECONDS = 7 * 60
MIN_SERMON_START_SECONDS = 15 * 60
MAX_SERMON_START_SECONDS = 45 * 60
MIN_SERMON_DURATION_SECONDS = 15 * 60
MAX_SERMON_DURATION_SECONDS = 50 * 60
PRE_START_CONTEXT_WINDOW_SECONDS = 60
POST_END_CONTEXT_WINDOW_SECONDS = 180
POST_APPLAUSE_LOOKAHEAD_SECONDS = 60
APPLAUSE_SEARCH_WINDOW_SECONDS = 20
START_CONTEXT_MARKERS = [
    "\ubc15\uc218",
    "\uc131\uac00\ub300",
    "\ucc2c\uc591\ub300",
    "\ucc2c\uc591",
    "\ud2b9\uc1a1",
]
CHOIR_CONTEXT_MARKERS = [
    "\uc131\uac00\ub300",
    "\ucc2c\uc591\ub300",
    "\ucc2c\uc591",
    "\ud2b9\uc1a1",
]
APPLAUSE_MARKERS = [
    "\ubc15\uc218",
]
START_SERMON_TEXT_MARKERS = [
    "\ud558\ub098\ub2d8\uc758 \ub9d0\uc500",
    "\ubcf8\ubb38",
    "\ub9d0\uc500",
    "\uc131\uacbd",
]
START_INTRO_MARKERS = [
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4",
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4 \uc5ec\ub7ec\ubd84",
    "\uc624\ub298 \uc6b0\ub9ac\ub294",
    "\uc624\ub298\uc740",
    "\uc624\ub298 \ub9d0\uc500\uc740",
    "\ud568\uaed8 \ubcfc \ub9d0\uc500",
    "\ubcf8\ubb38 \ub9d0\uc500",
    "\ubcf8\ubb38\uc740",
]
PRE_PRAYER_TRANSITION_MARKERS = [
    "\uc544\uba58",
    "\ub2e4\ud568\uaed8",
]
PRE_SERMON_SEQUENCE_MARKERS = [
    "\uc624\ub298 \uc6b0\ub9ac\ub294",
    "\uc624\ub298\uc740",
    "\uc624\ub298 \ub9d0\uc500\uc740",
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4",
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4 \uc5ec\ub7ec\ubd84",
    "\ub9d0\uc500\uc744 \uc804\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\ud568\uaed8 \ubcfc \ub9d0\uc500",
    "\ubcf8\ubb38 \ub9d0\uc500",
    "\ubcf8\ubb38\uc740",
    "\uace0\uac1c \uc219\uc5ec \uae30\ub3c4 \ud558\uaca0\uc2b5\ub2c8\ub2e4",
]
STRONG_SERMON_FOLLOWUP_MARKERS = [
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4",
    "\uc0ac\ub791\ud558\ub294 \uc131\ub3c4 \uc5ec\ub7ec\ubd84",
    "\uc624\ub298 \uc6b0\ub9ac\ub294",
    "\uc624\ub298\uc740",
    "\uc624\ub298 \ub9d0\uc500\uc740",
    "\ub9d0\uc500\uc744 \uc804\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\ud568\uaed8 \ubcfc \ub9d0\uc500",
    "\ubcf8\ubb38 \ub9d0\uc500",
    "\ubcf8\ubb38\uc740",
    "\uace0\uac1c \uc219\uc5ec \uae30\ub3c4 \ud558\uaca0\uc2b5\ub2c8\ub2e4",
]
SCRIPTURE_READING_MARKERS = [
    "\uc0ac\ub3c4\ud589\uc804",
    "\ub9c8\ud0dc\ubcf5\uc74c",
    "\ub9c8\uac00\ubcf5\uc74c",
    "\ub204\uac00\ubcf5\uc74c",
    "\uc694\ud55c\ubcf5\uc74c",
    "\ub85c\ub9c8\uc11c",
    "\uace0\ub9b0\ub3c4",
    "\uac08\ub77c\ub514\uc544\uc11c",
    "\uc5d0\ubca0\uc18c\uc11c",
    "\ube4c\ub9bd\ubcf4\uc11c",
    "\uace8\ub85c\uc0c8\uc11c",
    "\ub370\uc0b4\ub85c\ub2c8\uac00",
    "\ub514\ubaa8\ub370",
    "\ub514\ub3c4\uc11c",
    "\ud788\ube0c\ub9ac\uc11c",
    "\uc57c\uace0\ubcf4\uc11c",
    "\ubca0\ub4dc\ub85c",
    "\uc694\ud55c\uacc4\uc2dc\ub85d",
]
PRAYER_CONTINUATION_MARKERS = [
    "\uae30\ub3c4\ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4 \ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
    "\uae30\ub3c4\ud558\uc2dc\uaca0\uc2b5\ub2c8\ub2e4",
    "\uc608\uc218\ub2d8 \uc774\ub984\uc73c\ub85c",
    "\uc544\uba58",
]
OPENING_PRAYER_CONTENT_MARKERS = [
    "\ud558\ub098\ub2d8 \uc544\ubc84\uc9c0 \uac10\uc0ac\ud569\ub2c8\ub2e4",
    "\ud558\ub098\ub2d8\uc544\ubc84\uc9c0 \uac10\uc0ac\ud569\ub2c8\ub2e4",
    "\uc544\ubc84\uc9c0 \ud558\ub098\ub2d8 \uac10\uc0ac\ud569\ub2c8\ub2e4",
    "\uc544\ubc84\uc9c0 \uac10\uc0ac\ud569\ub2c8\ub2e4",
    "\uc740\ud61c\uc758 \ud558\ub098\ub2d8",
]
OPENING_PRAYER_FOLLOWUP_MARKERS = [
    "\uc624\ub298 \ub9d0\uc500",
    "\uc624\ub298 \ub9d0\uc500\uc744 \ub4dc\ub9b4 \ub54c",
    "\ub9d0\uc500\uc744 \ub4dc\ub9b4 \ub54c",
    "\ub9d0\uc500\uc744 \ud1b5\ud574\uc11c",
    "\uc131\uacbd\uc758 \ub9d0\uc500\uc744 \ud1b5\ud574\uc11c",
]
END_PRAYER_MARKERS = [
    "\uc608\uc218\ub2d8 \uc774\ub984\uc73c\ub85c \uae30\ub3c4 \ub4dc\ub838\uc2b5\ub2c8\ub2e4",
    "\uc608\uc218\ub2d8 \uc774\ub984\uc73c\ub85c \uae30\ub3c4\ud588\uc2b5\ub2c8\ub2e4",
    "\uc608\uc218\ub2d8 \uc774\ub984\uc73c\ub85c \uae30\ub3c4 \ub4dc\ub9bd\ub2c8\ub2e4",
    "\uc544\uba58",
]


def normalize_search_text(value: str) -> str:
    return re.sub(r"[\s\W_]+", "", value, flags=re.UNICODE).lower()


def contains_marker(text: str, markers: list[str]) -> bool:
    normalized_text = normalize_search_text(text)
    return any(normalize_search_text(marker) in normalized_text for marker in markers)


def parse_ffmpeg_progress_seconds(line: str) -> float | None:
    if line.startswith("out_time_ms="):
        try:
            return max(0.0, float(line.split("=", 1)[1]) / 1_000_000.0)
        except ValueError:
            return None
    if line.startswith("out_time_us="):
        try:
            return max(0.0, float(line.split("=", 1)[1]) / 1_000_000.0)
        except ValueError:
            return None
    if line.startswith("out_time="):
        raw_value = line.split("=", 1)[1].strip()
        try:
            return parse_timecode(raw_value)
        except ValueError:
            return None
    return None


def slugify(value: str) -> str:
    slug = re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE).strip().lower()
    slug = re.sub(r"[-\s]+", "-", slug)
    return slug or "clip"


def extract_youtube_id(url: str) -> str:
    parsed = urlparse(url.strip())
    host = parsed.netloc.lower()
    if "youtu.be" in host:
        video_id = parsed.path.strip("/")
        if video_id:
            return video_id
    if "youtube.com" in host or "music.youtube.com" in host:
        query_id = parse_qs(parsed.query).get("v", [""])[0].strip()
        if query_id:
            return query_id
        path_parts = [part for part in parsed.path.split("/") if part]
        if len(path_parts) >= 2 and path_parts[0] in {"shorts", "live", "embed"}:
            return path_parts[1]
    return ""


def build_job_slug(url: str, title: str = "") -> str:
    title_slug = slugify(title)
    if title_slug and title_slug != "clip":
        return title_slug
    video_id = extract_youtube_id(url)
    if video_id:
        return f"youtube-{slugify(video_id)}"
    digest = hashlib.sha1(url.strip().encode("utf-8")).hexdigest()[:10]
    return f"youtube-{digest}"


def parse_timecode(raw: str) -> float:
    parts = raw.strip().split(":")
    if not 1 <= len(parts) <= 3:
        raise ValueError(f"Invalid timecode: {raw}")
    values = [float(part) for part in parts]
    while len(values) < 3:
        values.insert(0, 0.0)
    hours, minutes, seconds = values
    return hours * 3600 + minutes * 60 + seconds


def format_timestamp(seconds: float) -> str:
    total_ms = round(seconds * 1000)
    hours, remainder = divmod(total_ms, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    secs, millis = divmod(remainder, 1000)
    return f"{hours:02}:{minutes:02}:{secs:02}.{millis:03}"


def resolve_binary(name: str, explicit: str = "") -> str:
    if explicit and Path(explicit).exists():
        return explicit

    bundled_candidates = [
        RUNTIME_DIR / f"{name}.exe",
        RUNTIME_DIR / "tools" / f"{name}.exe",
        PROJECT_DIR / "tools" / f"{name}.exe",
    ]
    for candidate in bundled_candidates:
        if candidate.exists():
            return str(candidate)

    discovered = shutil.which(name)
    if discovered:
        return discovered

    if name == "ffmpeg":
        ffmpeg_candidates = list(
            Path.home().glob(
                "AppData/Local/Microsoft/WinGet/Packages/Gyan.FFmpeg*/**/bin/ffmpeg.exe"
            )
        )
        if ffmpeg_candidates:
            return str(ffmpeg_candidates[0])

    raise FileNotFoundError(
        f"Could not find '{name}'. Install it or set the path in Settings."
    )


def resolve_ffprobe(ffmpeg_path: str) -> str:
    sibling = Path(ffmpeg_path).with_name("ffprobe.exe")
    if sibling.exists():
        return str(sibling)

    discovered = shutil.which("ffprobe")
    if discovered:
        return discovered

    raise FileNotFoundError("Could not find ffprobe.exe next to ffmpeg.exe.")


def run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW
    try:
        return subprocess.run(
            command,
            check=True,
            text=True,
            capture_output=True,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except subprocess.CalledProcessError as exc:
        details = []
        if exc.stdout:
            details.append(exc.stdout.strip())
        if exc.stderr:
            details.append(exc.stderr.strip())
        if details:
            raise RuntimeError("\n\n".join(part for part in details if part)) from exc
        raise


def run_command_streaming(
    command: list[str],
    line_handler=None,
) -> subprocess.CompletedProcess[str]:
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        startupinfo=startupinfo,
        creationflags=creationflags,
    )
    lines: list[str] = []
    assert process.stdout is not None
    for raw_line in process.stdout:
        line = raw_line.rstrip()
        lines.append(line)
        if line_handler:
            line_handler(line)
    return_code = process.wait()
    output = "\n".join(lines).strip()
    if return_code != 0:
        raise RuntimeError(output or f"Command failed with exit code {return_code}")
    return subprocess.CompletedProcess(command, return_code, stdout=output, stderr="")


@dataclass
class DownloadResult:
    source_file: Path
    title_slug: str


@dataclass
class YouTubeLiveCandidate:
    video_id: str
    title: str
    actual_end_time: str
    processing_status: str

    @property
    def url(self) -> str:
        return f"https://www.youtube.com/watch?v={self.video_id}"


@dataclass
class StartCandidate:
    score: int
    time_seconds: float
    reason: str


@dataclass
class VisionFrameResult:
    time_seconds: float
    score: int
    label: str
    reason: str


class SermonStudioEngine:
    def __init__(self, log_callback):
        self.log = log_callback
        self.status_callback = lambda message: None
        self.progress_callback = lambda value, maximum=None: None
        self.settings = self._load_settings()
        for path in [
            WORK_DIR,
            DOWNLOADS_DIR,
            EXPORTS_DIR,
            PREVIEW_DIR,
            TRANSCRIPT_DIR,
            FRAME_DIR,
            REVIEW_DIR,
        ]:
            path.mkdir(parents=True, exist_ok=True)

    def _load_settings(self) -> dict[str, Any]:
        if SETTINGS_PATH.exists():
            try:
                settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
                if isinstance(settings, dict):
                    settings.pop("openai_api_key", None)
                    return settings
                return {}
            except json.JSONDecodeError:
                return {}
        return {}

    def save_settings(self, payload: dict[str, Any]) -> None:
        payload.pop("openai_api_key", None)
        self.settings.pop("openai_api_key", None)
        self.settings.update(payload)
        SETTINGS_PATH.write_text(
            json.dumps(self.settings, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def get_setting(self, key: str, default: str = "") -> str:
        if key == "openai_api_key":
            if os.name == "nt":
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key_handle:
                        user_value, _ = winreg.QueryValueEx(key_handle, "OPENAI_API_KEY")
                    return user_value or default
                except OSError:
                    pass
            process_value = os.environ.get("OPENAI_API_KEY")
            if process_value:
                return process_value
            return default
        value = self.settings.get(key, default)
        return value if isinstance(value, str) else default

    def save_openai_api_key(self, api_key: str) -> None:
        api_key = api_key.strip()
        if not api_key:
            return
        os.environ["OPENAI_API_KEY"] = api_key
        if os.name == "nt":
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as key_handle:
                winreg.SetValueEx(key_handle, "OPENAI_API_KEY", 0, winreg.REG_SZ, api_key)
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key_handle:
                saved_value, _ = winreg.QueryValueEx(key_handle, "OPENAI_API_KEY")
            if saved_value != api_key:
                raise RuntimeError("OPENAI_API_KEY was not saved to Windows user environment.")

    def _parse_youtube_time(self, value: str) -> datetime:
        if not value:
            return datetime.min.replace(tzinfo=timezone.utc)
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return datetime.min.replace(tzinfo=timezone.utc)

    def get_youtube_service(self):
        client_secrets_path = self.get_setting("youtube_client_secrets_path")
        if not client_secrets_path:
            raise ValueError(
                "YouTube 계정 연결이 필요합니다.\n\n"
                "설정에서 Google OAuth client secrets JSON 파일 경로를 먼저 지정해 주세요."
            )
        if not Path(client_secrets_path).exists():
            raise FileNotFoundError(f"Google OAuth client secrets JSON 파일을 찾을 수 없습니다:\n{client_secrets_path}")

        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "YouTube API 라이브러리가 설치되어 있지 않습니다. "
                "패키지를 다시 빌드하거나 requirements를 설치해야 합니다."
            ) from exc

        scopes = ["https://www.googleapis.com/auth/youtube.readonly"]
        credentials = None
        if YOUTUBE_TOKEN_PATH.exists():
            credentials = Credentials.from_authorized_user_file(str(YOUTUBE_TOKEN_PATH), scopes)
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            self.status("Opening Google login for YouTube access...")
            flow = InstalledAppFlow.from_client_secrets_file(str(client_secrets_path), scopes)
            credentials = flow.run_local_server(port=0, prompt="consent")
        YOUTUBE_TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        YOUTUBE_TOKEN_PATH.write_text(credentials.to_json(), encoding="utf-8")
        return build("youtube", "v3", credentials=credentials)

    def find_recent_live_archive(self, service_order: str) -> YouTubeLiveCandidate:
        if service_order not in {"first", "second"}:
            raise ValueError("service_order must be 'first' or 'second'.")

        service = self.get_youtube_service()
        label = "1부" if service_order == "first" else "2부"
        self.status(f"Finding recent completed live archive for {label}...")
        self.progress(None)

        response = service.liveBroadcasts().list(
            part="id,snippet,status",
            mine=True,
            broadcastStatus="completed",
            maxResults=10,
        ).execute()
        items = response.get("items", [])
        if len(items) < 1:
            raise RuntimeError("완료된 라이브 방송을 찾지 못했습니다.")

        def sort_key(item: dict[str, Any]) -> datetime:
            snippet = item.get("snippet", {})
            return self._parse_youtube_time(
                snippet.get("actualEndTime")
                or snippet.get("actualStartTime")
                or snippet.get("scheduledStartTime")
                or snippet.get("publishedAt")
                or ""
            )

        items.sort(key=sort_key, reverse=True)
        index = 1 if service_order == "first" else 0
        if len(items) <= index:
            raise RuntimeError(f"{label} 예배 후보를 찾기에는 완료된 라이브가 부족합니다.")

        selected = items[index]
        video_id = str(selected.get("id", "")).strip()
        snippet = selected.get("snippet", {})
        title = str(snippet.get("title", "")).strip()
        actual_end_time = str(
            snippet.get("actualEndTime")
            or snippet.get("actualStartTime")
            or snippet.get("scheduledStartTime")
            or ""
        )
        if not video_id:
            raise RuntimeError("선택된 라이브에서 YouTube video ID를 찾지 못했습니다.")

        processing_status = "unknown"
        try:
            video_response = service.videos().list(
                part="processingDetails,status,snippet",
                id=video_id,
            ).execute()
            video_items = video_response.get("items", [])
            if video_items:
                processing_status = str(
                    video_items[0].get("processingDetails", {}).get("processingStatus", "unknown")
                )
        except Exception as exc:
            self.log(f"Could not read YouTube processing status: {exc}")

        candidate = YouTubeLiveCandidate(
            video_id=video_id,
            title=title,
            actual_end_time=actual_end_time,
            processing_status=processing_status,
        )
        self.log(
            f"{label} live archive candidate: {candidate.url} | "
            f"title={candidate.title} | processing={candidate.processing_status}"
        )
        self.progress(100, 100)
        return candidate

    def set_callbacks(self, log_callback, status_callback, progress_callback) -> None:
        self.log = log_callback
        self.status_callback = status_callback
        self.progress_callback = progress_callback

    def status(self, message: str) -> None:
        self.status_callback(message)

    def progress(self, value: int | None, maximum: int | None = None) -> None:
        self.progress_callback(value, maximum)

    def download_video(
        self,
        url: str,
        title: str,
        yt_dlp_path: str,
    ) -> DownloadResult:
        if not url:
            raise ValueError("Enter a YouTube URL first.")

        yt_dlp_bin = resolve_binary("yt-dlp", yt_dlp_path)
        title_slug = build_job_slug(url, title)
        target_dir = DOWNLOADS_DIR / title_slug
        target_dir.mkdir(parents=True, exist_ok=True)

        existing = sorted(
            path for path in target_dir.glob("source.*")
            if not path.name.endswith((".part", ".ytdl", ".temp"))
        )
        if existing:
            self.log(f"Reusing existing download: {existing[0]}")
            self.status("Using existing downloaded video.")
            self.progress(100, 100)
            return DownloadResult(source_file=existing[0], title_slug=title_slug)

        output_template = str(target_dir / "source.%(ext)s")
        self.status("Downloading full service video...")
        self.progress(0, 100)

        def handle_download_line(line: str) -> None:
            if not line:
                return
            self.log(line)
            match = re.search(r"(\d+(?:\.\d+)?)%", line)
            if match:
                self.progress(min(100, int(float(match.group(1)))), 100)

        def build_download_command(browser: str = "") -> list[str]:
            command = [
                yt_dlp_bin,
                "--newline",
                "--ignore-config",
                "--no-playlist",
                "--continue",
                "--retries",
                "10",
                "--fragment-retries",
                "10",
                "--retry-sleep",
                "fragment:2",
                "--extractor-args",
                "youtube:player_client=android,web",
                "-f",
                "bv*[ext=mp4][vcodec^=avc1][height<=1080][fps<=60]+ba[ext=m4a]/b[ext=mp4][height<=1080]/b",
                "-S",
                "res:1080,fps,vcodec:h264,acodec:m4a",
                "--merge-output-format",
                "mp4",
            ]
            if browser:
                command.extend(["--cookies-from-browser", browser])
            command.extend(["-o", output_template, url])
            return command

        def needs_browser_cookies(error_text: str) -> bool:
            lowered = error_text.lower()
            return (
                "sign in to confirm" in lowered
                or "not a bot" in lowered
                or "--cookies-from-browser" in lowered
                or "cookies" in lowered
            )

        try:
            completed = run_command_streaming(build_download_command(), handle_download_line)
        except RuntimeError as exc:
            if not needs_browser_cookies(str(exc)):
                raise
            last_error = exc
            for browser in ["chrome", "edge", "firefox"]:
                self.log(f"Download needs YouTube login cookies. Retrying with {browser} browser cookies...")
                try:
                    completed = run_command_streaming(build_download_command(browser), handle_download_line)
                    break
                except RuntimeError as retry_exc:
                    last_error = retry_exc
                    self.log(f"Retry with {browser} cookies failed: {retry_exc}")
            else:
                last_error_text = str(last_error)
                if "could not copy chrome cookie database" in last_error_text.lower():
                    raise RuntimeError(
                        "YouTube 로그인 쿠키가 필요하지만, Chrome/Edge가 쿠키 파일을 사용 중이라 복사하지 못했습니다.\n\n"
                        "해결 방법:\n"
                        "1. Chrome과 Edge 창을 모두 닫습니다.\n"
                        "2. 작업 관리자에서 chrome.exe 또는 msedge.exe가 남아 있으면 종료합니다.\n"
                        "3. 다시 프로그램에서 다운로드를 실행합니다.\n\n"
                        "브라우저에 YouTube 로그인이 되어 있어도 브라우저가 켜져 있으면 이 에러가 날 수 있습니다.\n\n"
                        f"Last yt-dlp error:\n{last_error}"
                    ) from last_error
                raise RuntimeError(
                    "YouTube requires browser login cookies for this video. "
                    "Open the video in Chrome or Edge while signed in, close the browser completely, then try again.\n\n"
                    f"Last yt-dlp error:\n{last_error}"
                ) from last_error
        if completed.stdout:
            self.log(completed.stdout.strip())

        downloaded = sorted(
            path for path in target_dir.glob("source.*")
            if not path.name.endswith((".part", ".ytdl", ".temp"))
        )
        if not downloaded:
            raise FileNotFoundError("Download finished but no source file was found.")
        self.status("Download completed.")
        self.progress(100, 100)
        return DownloadResult(source_file=downloaded[0], title_slug=title_slug)

    def extract_full_audio(self, source_file: Path, title_slug: str, ffmpeg_path: str) -> Path:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        audio_path = PREVIEW_DIR / f"{title_slug}-full.mp3"
        command = [
            ffmpeg_bin,
            "-y",
            "-i",
            str(source_file),
            "-vn",
            "-ac",
            "1",
            "-ar",
            "24000",
            "-c:a",
            "libmp3lame",
            "-b:a",
            "64k",
            str(audio_path),
        ]
        self.status("Extracting full audio for transcription...")
        self.progress(None)
        self.log("Extracting full audio for transcription...")
        run_command(command)
        return audio_path

    def split_audio_chunks(
        self,
        audio_file: Path,
        title_slug: str,
        ffmpeg_path: str,
        chunk_minutes: int,
        skip_start_seconds: int = DEFAULT_TRANSCRIPTION_SKIP_START_SECONDS,
        skip_end_seconds: int = DEFAULT_TRANSCRIPTION_SKIP_END_SECONDS,
    ) -> list[tuple[Path, float]]:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        chunk_dir = PREVIEW_DIR / f"{title_slug}-chunks"
        chunk_dir.mkdir(parents=True, exist_ok=True)
        for old_chunk in chunk_dir.glob("chunk-*.*"):
            old_chunk.unlink()

        chunk_seconds = max(60, chunk_minutes * 60)
        ffprobe_bin = resolve_ffprobe(ffmpeg_bin)
        duration_command = [
            ffprobe_bin,
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(audio_file),
        ]
        duration_result = run_command(duration_command)
        total_seconds = float(duration_result.stdout.strip())
        analysis_start = min(max(0, skip_start_seconds), max(0, int(total_seconds) - 60))
        analysis_end = max(analysis_start + 60, int(total_seconds) - max(0, skip_end_seconds))
        analysis_duration = max(60, analysis_end - analysis_start)

        command = [
            ffmpeg_bin,
            "-y",
            "-ss",
            format_timestamp(analysis_start),
            "-i",
            str(audio_file),
            "-t",
            str(analysis_duration),
            "-f",
            "segment",
            "-segment_time",
            str(chunk_seconds),
            "-c:a",
            "copy",
            str(chunk_dir / "chunk-%03d.mp3"),
        ]
        self.status(f"Splitting audio into {chunk_minutes}-minute chunks...")
        self.progress(None)
        self.log(
            "Splitting audio into "
            f"{chunk_minutes}-minute chunks after skipping the first {skip_start_seconds // 60} minutes "
            f"and the last {skip_end_seconds // 60} minutes."
        )
        run_command(command)

        chunks = sorted(chunk_dir.glob("chunk-*.mp3"))
        if not chunks:
            raise FileNotFoundError("No audio chunks were created.")
        return [
            (chunk, analysis_start + index * float(chunk_seconds))
            for index, chunk in enumerate(chunks)
        ]

    def transcribe_chunks(
        self,
        transcript_name: str,
        model_name: str,
        api_key: str,
        language: str,
        chunks: list[tuple[Path, float]],
    ) -> Path:
        if not api_key:
            raise ValueError("OpenAI API key is missing. Add it in Settings.")

        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        transcript_path = TRANSCRIPT_DIR / f"{transcript_name}.json"
        combined_segments: list[dict[str, Any]] = []
        combined_text_parts: list[str] = []
        response_format = "diarized_json" if model_name == "gpt-4o-transcribe-diarize" else "verbose_json"

        total_chunks = max(1, len(chunks))
        for chunk_index, (chunk_file, offset_seconds) in enumerate(chunks, start=1):
            self.status(f"Transcribing chunk {chunk_index} of {total_chunks}...")
            self.progress(int(((chunk_index - 1) / total_chunks) * 100), 100)
            self.log(f"Transcribing chunk: {chunk_file.name}")
            with chunk_file.open("rb") as audio_file:
                kwargs: dict[str, Any] = {
                    "model": model_name,
                    "file": audio_file,
                    "language": language or "ko",
                    "response_format": response_format,
                }
                if model_name == "gpt-4o-transcribe-diarize":
                    kwargs["chunking_strategy"] = "auto"
                else:
                    kwargs["prompt"] = (
                        "This is Korean church worship audio. "
                        "Transcribe transitions into the sermon, scripture reading, prayer, "
                        "and announcements as clearly as possible."
                    )
                    if model_name == "whisper-1":
                        kwargs["timestamp_granularities"] = ["segment", "word"]
                transcript = client.audio.transcriptions.create(**kwargs)

            if hasattr(transcript, "model_dump"):
                payload = transcript.model_dump()
            elif isinstance(transcript, dict):
                payload = transcript
            else:
                payload = {"text": getattr(transcript, "text", "")}

            combined_text_parts.append(str(payload.get("text", "")).strip())
            for index, segment in enumerate(payload.get("segments") or []):
                start = float(segment.get("start", 0.0)) + offset_seconds
                end = float(segment.get("end", start)) + offset_seconds
                combined_segments.append(
                    {
                        "id": segment.get("id", f"seg-{len(combined_segments)+1}"),
                        "start": start,
                        "end": end,
                        "text": str(segment.get("text", "")).strip(),
                        "speaker": segment.get("speaker"),
                        "chunk_index": index,
                    }
                )

            self.progress(int((chunk_index / total_chunks) * 100), 100)

        payload = {
            "model": model_name,
            "text": "\n".join(part for part in combined_text_parts if part),
            "segments": combined_segments,
        }
        transcript_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return transcript_path

    def detect_applause_end(
        self,
        audio_file: Path,
        ffmpeg_path: str,
        window_start_seconds: float,
        window_end_seconds: float,
    ) -> float | None:
        if window_end_seconds <= window_start_seconds:
            return None

        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        clip_duration = max(2.0, window_end_seconds - window_start_seconds)
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_file:
            temp_wav = Path(temp_file.name)

        try:
            command = [
                ffmpeg_bin,
                "-y",
                "-ss",
                format_timestamp(window_start_seconds),
                "-i",
                str(audio_file),
                "-t",
                str(clip_duration),
                "-ac",
                "1",
                "-ar",
                "8000",
                "-c:a",
                "pcm_s16le",
                str(temp_wav),
            ]
            run_command(command)

            with wave.open(str(temp_wav), "rb") as wav_file:
                sample_rate = wav_file.getframerate()
                sample_width = wav_file.getsampwidth()
                frame_samples = max(1, int(sample_rate * 0.04))
                rms_values: list[int] = []
                while True:
                    frames = wav_file.readframes(frame_samples)
                    if not frames:
                        break
                    rms_values.append(audioop.rms(frames, sample_width))

            if len(rms_values) < 20:
                return None

            baseline = max(1.0, statistics.median(rms_values))
            upper_band = max(baseline * 1.6, statistics.quantiles(rms_values, n=10)[7])
            high_frames = [index for index, value in enumerate(rms_values) if value >= upper_band]
            if not high_frames:
                return None

            clusters: list[tuple[int, int]] = []
            cluster_start = high_frames[0]
            cluster_end = high_frames[0]
            for index in high_frames[1:]:
                if index - cluster_end <= 3:
                    cluster_end = index
                else:
                    clusters.append((cluster_start, cluster_end))
                    cluster_start = index
                    cluster_end = index
            clusters.append((cluster_start, cluster_end))

            best_cluster: tuple[int, int] | None = None
            best_score = -1.0
            total_frames = len(rms_values)
            for cluster in clusters:
                start_frame, end_frame = cluster
                duration_frames = end_frame - start_frame + 1
                if duration_frames < 8:
                    continue
                # Prefer the last dense burst shortly before the prayer segment.
                recency_bonus = 1.0 - ((total_frames - end_frame) / total_frames)
                score = duration_frames + (recency_bonus * 20.0)
                if score > best_score:
                    best_score = score
                    best_cluster = cluster

            if not best_cluster:
                return None

            _, cluster_end = best_cluster
            frame_seconds = 0.04
            applause_end = window_start_seconds + ((cluster_end + 1) * frame_seconds)
            if applause_end >= window_end_seconds:
                applause_end = max(window_start_seconds, window_end_seconds - 0.2)
            return applause_end
        finally:
            temp_wav.unlink(missing_ok=True)

    def extract_frame_at_time(
        self,
        source_file: Path,
        ffmpeg_path: str,
        capture_seconds: float,
        frame_name: str,
    ) -> Path:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        frame_path = FRAME_DIR / frame_name
        command = [
            ffmpeg_bin,
            "-y",
            "-ss",
            format_timestamp(max(0.0, capture_seconds)),
            "-i",
            str(source_file),
            "-frames:v",
            "1",
            "-q:v",
            "2",
            str(frame_path),
        ]
        run_command(command)
        return frame_path

    def score_start_frame_with_vision(
        self,
        frame_path: Path,
        api_key: str,
        vision_model: str,
    ) -> tuple[int, str, str]:
        if not api_key:
            return 0, "unknown", "no api key"

        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        image_data = base64.b64encode(frame_path.read_bytes()).decode("ascii")
        response = client.responses.create(
            model=vision_model or "gpt-4.1-mini",
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "This is a frame from a Korean church worship video. "
                "Classify this frame as one of: sermon_title_slide, sermon_start, scripture_reading, choir_or_praise, prayer_slide_or_prayer, other. "
                "The ONLY best sermon start marker is the first sermon title slide immediately after choir/praise choir. "
                "Classify as sermon_title_slide only when the visible slide has the FGMC sermon format: "
                "the word '말씀' near the top, a sermon title, a Bible reference such as '사도행전 1장 1절' or '행 1:15-26', "
                "and a pastor/preacher name with a role such as '김일영 목사', '담임목사', '선교사', or another Korean name plus 목사/선교사. "
                "If this frame appears right after choir and is a title-card style slide with '말씀', treat it as the sermon title slide "
                "even if some small text is hard to read. "
                "If the frame is a pastor at the pulpit without that sermon-title slide, do not call it sermon_title_slide. "
                "Treat visible on-screen text like '성경봉독', plain Bible passage reading slides, or scripture-reading lower thirds "
                "as scripture_reading, not sermon_title_slide and not sermon_start. "
                "Do NOT classify prayer slides such as '합심기도', '대표기도', worship order slides, announcements, or generic prayer slides "
                "as sermon_title_slide. "
                "Treat choir, praise team, congregation singing, or music performance visuals "
                "as choir_or_praise. "
                                "Treat prayer slides like '합심기도', visible prayer-only title slides, or early representative prayer scenes "
                                "before the sermon as prayer_slide_or_prayer, not sermon_start. "
                                "Read the visible Korean/English text carefully. "
                                "Score from 0 to 10 how likely this exact frame is the FGMC sermon title slide. "
                                "Return JSON only like "
                                "{\"label\":\"sermon_title_slide\",\"score\":9,\"visible_text\":\"말씀 ... 홍길동 목사\",\"reason\":\"short phrase\"}."
                            ),
                        },
                        {
                            "type": "input_image",
                            "image_url": f"data:image/jpeg;base64,{image_data}",
                            "detail": "high",
                        },
                    ],
                }
            ],
        )
        output = getattr(response, "output_text", "").strip()
        label_match = re.search(r'"label"\s*:\s*"([^"]+)"', output)
        match = re.search(r'"score"\s*:\s*(\d+)', output)
        reason_match = re.search(r'"reason"\s*:\s*"([^"]*)"', output)
        visible_text_match = re.search(r'"visible_text"\s*:\s*"([^"]*)"', output)
        label = (label_match.group(1).strip().lower() if label_match else "unknown")
        visible_text = visible_text_match.group(1).strip() if visible_text_match else ""
        reason = reason_match.group(1).strip() if reason_match else output[:120]
        if not match:
            return 0, label, reason
        score = max(0, min(10, int(match.group(1))))

        normalized_visible_text = normalize_search_text(visible_text)
        excluded_slide_markers = [
            "\ud569\uc2ec\uae30\ub3c4",
            "\ub300\ud45c\uae30\ub3c4",
            "\uc131\uacbd\ubd09\ub3c5",
        ]
        has_excluded_slide_text = any(
            normalize_search_text(marker) in normalized_visible_text
            for marker in excluded_slide_markers
        )
        has_sermon_title_text = (
            "\ub9d0\uc500" in normalized_visible_text
            and (
                "\ubaa9\uc0ac" in normalized_visible_text
                or "\ub2f4\uc784\ubaa9\uc0ac" in normalized_visible_text
                or "\uc120\uad50\uc0ac" in normalized_visible_text
            )
            and (
                "\uc0ac\ub3c4\ud589\uc804" in normalized_visible_text
                or re.search("행\\d", normalized_visible_text) is not None
                or re.search(r"\d{3,}", normalized_visible_text) is not None
                or ("\uc7a5" in normalized_visible_text and "\uc808" in normalized_visible_text)
            )
        )
        if has_excluded_slide_text:
            return min(score, 2), "prayer_slide_or_prayer" if "\uae30\ub3c4" in normalized_visible_text else "scripture_reading", f"{reason}; visible_text={visible_text}"
        if has_sermon_title_text:
            return max(score, 9), "sermon_title_slide", f"{reason}; visible_text={visible_text}"
        return score, label, f"{reason}; visible_text={visible_text}" if visible_text else reason

    def classify_start_frame(
        self,
        source_file: Path,
        ffmpeg_path: str,
        api_key: str,
        vision_model: str,
        title_slug: str,
        capture_seconds: float,
        suffix: str,
    ) -> VisionFrameResult:
        frame_name = f"{title_slug}-{suffix}-{int(capture_seconds * 1000)}.jpg"
        frame_path = self.extract_frame_at_time(
            source_file=source_file,
            ffmpeg_path=ffmpeg_path,
            capture_seconds=capture_seconds,
            frame_name=frame_name,
        )
        vision_score, vision_label, vision_reason = self.score_start_frame_with_vision(
            frame_path,
            api_key,
            vision_model,
        )
        return VisionFrameResult(
            time_seconds=capture_seconds,
            score=vision_score,
            label=vision_label,
            reason=vision_reason,
        )

    def extract_start_contact_sheet(
        self,
        source_file: Path,
        ffmpeg_path: str,
        title_slug: str,
        start_seconds: float,
        duration_seconds: float,
        sample_step_seconds: int,
        columns: int,
        rows: int,
        suffix: str,
    ) -> Path:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        frame_path = FRAME_DIR / f"{title_slug}-{suffix}.jpg"
        vf = f"fps=1/{sample_step_seconds},scale=320:-1,tile={columns}x{rows}"
        command = [
            ffmpeg_bin,
            "-y",
            "-ss",
            format_timestamp(max(0.0, start_seconds)),
            "-t",
            format_timestamp(duration_seconds),
            "-i",
            str(source_file),
            "-frames:v",
            "1",
            "-vf",
            vf,
            "-q:v",
            "3",
            str(frame_path),
        ]
        run_command(command)
        return frame_path

    def find_sermon_transition_from_contact_sheet(
        self,
        contact_sheet_path: Path,
        api_key: str,
        vision_model: str,
        start_seconds: float,
        sample_step_seconds: int,
        max_tiles: int,
    ) -> tuple[float, float | None] | None:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        image_data = base64.b64encode(contact_sheet_path.read_bytes()).decode("ascii")
        response = client.responses.create(
            model=vision_model or "gpt-4.1-mini",
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "This contact sheet is from one Korean church worship video. "
                                "Tiles are in row-major order, starting at tile_index 0. "
                                f"Tile 0 is {format_timestamp(start_seconds)}, and each next tile is {sample_step_seconds} seconds later. "
                                "Find the LAST tile where the choir/praise choir is still visibly singing, "
                                "then find the FIRST sermon title slide immediately after that choir ends. "
                                "The correct slide format contains '말씀', a sermon title, a Bible reference, and a preacher name/role such as 목사, 담임목사, or 선교사. "
                                "The sermon title slide appears right after the choir ends, usually within a few seconds. "
                                "Do not choose 합심기도, 대표기도, 성경봉독, scripture-reading-only screens, announcements, or pastor-only frames. "
                                "Return JSON only like "
                                "{\"last_choir_tile\":31,\"sermon_title_tile\":32,\"score\":9,"
                                "\"visible_text\":\"말씀 ... 홍길동 목사\",\"reason\":\"title slide immediately after choir\"}. "
                                "If no matching transition exists, return {\"last_choir_tile\":-1,\"sermon_title_tile\":-1,\"score\":0,\"reason\":\"not found\"}."
                            ),
                        },
                        {
                            "type": "input_image",
                            "image_url": f"data:image/jpeg;base64,{image_data}",
                            "detail": "high",
                        },
                    ],
                }
            ],
        )
        output = getattr(response, "output_text", "").strip()
        choir_tile_match = re.search(r'"last_choir_tile"\s*:\s*(-?\d+)', output)
        title_tile_match = re.search(r'"sermon_title_tile"\s*:\s*(-?\d+)', output)
        score_match = re.search(r'"score"\s*:\s*(\d+)', output)
        reason_match = re.search(r'"reason"\s*:\s*"([^"]*)"', output)
        visible_text_match = re.search(r'"visible_text"\s*:\s*"([^"]*)"', output)
        if not choir_tile_match or not title_tile_match:
            self.log(f"Contact-sheet transition check returned unparseable output: {output[:200]}")
            return None
        choir_tile = int(choir_tile_match.group(1))
        title_tile = int(title_tile_match.group(1))
        score = max(0, min(10, int(score_match.group(1)))) if score_match else 0
        reason = reason_match.group(1).strip() if reason_match else output[:120]
        visible_text = visible_text_match.group(1).strip() if visible_text_match else ""
        choir_time = start_seconds + max(choir_tile, 0) * sample_step_seconds
        title_time = start_seconds + max(title_tile, 0) * sample_step_seconds
        self.log(
            f"Contact-sheet transition check: last_choir_tile={choir_tile} "
            f"choir_time={format_timestamp(choir_time)} sermon_title_tile={title_tile} "
            f"title_time={format_timestamp(title_time)} score={score} "
            f"detail={reason}; visible_text={visible_text}"
        )
        if (
            choir_tile < 0
            or title_tile < 0
            or choir_tile >= max_tiles
            or title_tile >= max_tiles
            or title_tile <= choir_tile
            or score < 5
        ):
            return None
        return choir_time, title_time

    def find_scene_based_start_candidate(
        self,
        source_file: Path,
        ffmpeg_path: str,
        api_key: str,
        vision_model: str,
        title_slug: str,
    ) -> StartCandidate | None:
        # Disabled for now: the previous scene/contact-sheet scans were slow and
        # could confuse choir frames with the sermon title slide. Start detection
        # now uses transcript candidates and only verifies nearby video frames.
        return None

    def refine_start_with_vision(
        self,
        source_file: Path,
        ffmpeg_path: str,
        api_key: str,
        vision_model: str,
        candidates: list[StartCandidate],
        title_slug: str,
    ) -> float | None:
        if not candidates or not api_key:
            return None

        sorted_candidates = sorted(candidates, key=lambda item: item.score, reverse=True)
        for candidate in sorted_candidates[:8]:
            if candidate.time_seconds < MIN_SERMON_START_SECONDS:
                continue
            if "prayer-content backoff" in candidate.reason:
                center_time = candidate.time_seconds + 14.0
            else:
                center_time = candidate.time_seconds
            window_start = max(MIN_SERMON_START_SECONDS, center_time - 45.0)
            window_end = min(MAX_SERMON_START_SECONDS, center_time + 90.0)
            current_time = window_start
            rejected_slide_hits = 0
            while current_time <= window_end:
                frame_name = f"{title_slug}-start-candidate-{int(candidate.time_seconds * 1000)}-{int(current_time * 1000)}.jpg"
                frame_path = self.extract_frame_at_time(
                    source_file=source_file,
                    ffmpeg_path=ffmpeg_path,
                    capture_seconds=current_time,
                    frame_name=frame_name,
                )
                vision_score, vision_label, vision_reason = self.score_start_frame_with_vision(
                    frame_path,
                    api_key,
                    vision_model,
                )
                self.log(
                    f"Vision start check: candidate={format_timestamp(candidate.time_seconds)} "
                    f"frame={format_timestamp(current_time)} base={candidate.score} "
                    f"vision={vision_score} label={vision_label} reason={candidate.reason} "
                    f"detail={vision_reason}"
                )
                if vision_label == "sermon_title_slide" and vision_score >= 5:
                    return current_time
                if vision_label in {"prayer_slide_or_prayer", "scripture_reading"} and vision_score >= 2:
                    rejected_slide_hits += 1
                    if rejected_slide_hits >= 2:
                        self.log(
                            f"Skipping start candidate {format_timestamp(candidate.time_seconds)} "
                            f"after repeated rejected slide label={vision_label}."
                        )
                        break
                current_time += 5.0

        return None

    def suggest_sermon_range(
        self,
        transcript_path: Path,
        audio_file: Path | None = None,
        ffmpeg_path: str = "",
        source_file: Path | None = None,
        api_key: str = "",
        vision_model: str = "",
        title_slug: str = "service-video",
    ) -> tuple[str, str]:
        payload = json.loads(transcript_path.read_text(encoding="utf-8"))
        segments = payload.get("segments") or []
        if not segments:
            raise ValueError("This transcript does not contain timestamped segments.")

        lowered_markers = [marker.lower() for marker in DEFAULT_MARKERS]
        lowered_end_markers = [marker.lower() for marker in DEFAULT_END_MARKERS]
        lowered_context_markers = [marker.lower() for marker in START_CONTEXT_MARKERS]
        lowered_choir_context_markers = [marker.lower() for marker in CHOIR_CONTEXT_MARKERS]
        lowered_applause_markers = [marker.lower() for marker in APPLAUSE_MARKERS]
        lowered_sermon_text_markers = [marker.lower() for marker in START_SERMON_TEXT_MARKERS]
        lowered_start_intro_markers = [marker.lower() for marker in START_INTRO_MARKERS]
        lowered_pre_sermon_sequence_markers = [marker.lower() for marker in PRE_SERMON_SEQUENCE_MARKERS]
        lowered_strong_sermon_followup_markers = [marker.lower() for marker in STRONG_SERMON_FOLLOWUP_MARKERS]
        lowered_scripture_reading_markers = [marker.lower() for marker in SCRIPTURE_READING_MARKERS]
        lowered_prayer_continuation_markers = [marker.lower() for marker in PRAYER_CONTINUATION_MARKERS]
        lowered_opening_prayer_content_markers = [marker.lower() for marker in OPENING_PRAYER_CONTENT_MARKERS]
        lowered_opening_prayer_followup_markers = [marker.lower() for marker in OPENING_PRAYER_FOLLOWUP_MARKERS]
        prayer_markers = [
            "\ub2e4\ud568\uaed8 \uae30\ub3c4 \ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
            "\ub2e4\ud568\uaed8 \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
            "\uace0\uac1c \uc219\uc5ec \uae30\ub3c4 \ud558\uaca0\uc2b5\ub2c8\ub2e4",
            "\uace0\uac1c\uc219\uc5ec \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
            "\uae30\ub3c4 \ub4dc\ub9ac\uaca0\uc2b5\ub2c8\ub2e4",
            "\uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
            "\uae30\ub3c4\ud558\uc2dc\uaca0\uc2b5\ub2c8\ub2e4",
        ]
        fuzzy_end_markers = [
            "\ucd95\uc6d0",
            "\ucd95\uc6d0\ud569\ub2c8\ub2e4",
            "\ucd95\uc6d0 \ub4dc\ub9bd\ub2c8\ub2e4",
            "\ucd95\uc6d0\ub4dc\ub9bd",
            "\ucd95\uc6d0\ud569",
            "\ucd94\uad8c",
            "\ucd94\uad8c\ud569\ub2c8\ub2e4",
            "\ucd94\uac74",
            "\ucd94\uac74\ud569\ub2c8\ub2e4",
            "\ucd95\uac74",
            "\ucd95\uac74\ud569\ub2c8\ub2e4",
            "\ucd94\uad6c",
            "\ucd94\uad6c\ud569\ub2c8\ub2e4",
        ]

        start_time = None
        end_time = None
        start_candidates: list[StartCandidate] = []
        if source_file and api_key:
            self.log("Start detection: using transcript candidates, then verifying nearby sermon title slide frames.")
        elif source_file and not api_key:
            self.log("Vision start verification skipped: OpenAI API key was not found in OPENAI_API_KEY.")

        for index, segment in enumerate(segments):
            text = str(segment.get("text", "")).strip().lower()
            segment_start = float(segment.get("start", 0.0))
            segment_end = float(segment.get("end", segment_start))

            if segment_start < MIN_SERMON_START_SECONDS:
                continue
            if segment_start > MAX_SERMON_START_SECONDS:
                break

            score = 0
            candidate_start = max(0.0, segment_start - 5.0)
            has_applause = any(marker in text for marker in lowered_applause_markers)
            has_prayer_start = contains_marker(text, prayer_markers)
            has_bow_prayer_start = contains_marker(
                text,
                [
                    "\uace0\uac1c \uc219\uc5ec \uae30\ub3c4 \ud558\uaca0\uc2b5\ub2c8\ub2e4",
                    "\uace0\uac1c\uc219\uc5ec \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
                ],
            )
            has_opening_prayer_content = contains_marker(text, OPENING_PRAYER_CONTENT_MARKERS)
            choir_context_window_start = max(MIN_SERMON_START_SECONDS, segment_start - 8 * 60.0)
            choir_context_segments = [
                prev for prev in segments
                if choir_context_window_start <= float(prev.get("start", 0.0)) < segment_start
            ]
            choir_context_text = " ".join(
                str(prev.get("text", "")).strip().lower()
                for prev in choir_context_segments
            )
            has_choir_context = any(marker in choir_context_text for marker in lowered_choir_context_markers)

            if has_prayer_start and audio_file and ffmpeg_path:
                detected_applause_end = self.detect_applause_end(
                    audio_file=audio_file,
                    ffmpeg_path=ffmpeg_path,
                    window_start_seconds=max(MIN_SERMON_START_SECONDS, segment_start - APPLAUSE_SEARCH_WINDOW_SECONDS),
                    window_end_seconds=segment_start,
                )
                if detected_applause_end is not None:
                    candidate_start = detected_applause_end
                    score += 12
                    candidate_reason = "audio applause end before prayer start"
                    self.log(
                        f"Audio applause assist: prayer marker at {format_timestamp(segment_start)} "
                        f"matched applause end near {format_timestamp(candidate_start)}."
                    )
                else:
                    candidate_reason = "prayer start without audio applause match"
            else:
                candidate_reason = "transcript heuristic"

            if has_opening_prayer_content and not has_bow_prayer_start:
                candidate_start = max(0.0, segment_start - 12.0)
                score += 8
                candidate_reason = f"{candidate_reason} + prayer-content backoff"

            if has_applause:
                candidate_start = segment_end
                score += 10

                follow_segments = []
                for next_segment in segments[index + 1:]:
                    next_start = float(next_segment.get("start", 0.0))
                    if next_start > candidate_start + POST_APPLAUSE_LOOKAHEAD_SECONDS:
                        break
                    follow_segments.append(next_segment)

                follow_text = " ".join(
                    str(next_segment.get("text", "")).strip().lower()
                    for next_segment in follow_segments
                )
                if contains_marker(follow_text, prayer_markers):
                    score += 6
                if contains_marker(follow_text, START_SERMON_TEXT_MARKERS):
                    score += 4
                if contains_marker(follow_text, DEFAULT_MARKERS):
                    score += 2
            else:
                if has_prayer_start:
                    score += 2
                if contains_marker(text, DEFAULT_MARKERS):
                    score += 2
                if contains_marker(text, START_SERMON_TEXT_MARKERS):
                    score += 2

            if has_prayer_start:
                score += 4
            if has_bow_prayer_start:
                score += 16
                candidate_reason = f"{candidate_reason} + bow prayer start"
            if has_opening_prayer_content:
                score += 4
            if has_choir_context and (has_applause or has_prayer_start or has_opening_prayer_content):
                score += 28
                candidate_reason = f"{candidate_reason} + after choir"
            elif has_applause or has_prayer_start or has_opening_prayer_content:
                score -= 24
                candidate_reason = f"{candidate_reason} - no choir context"

            # Prefer the first plausible sermon transition rather than later prayer moments.
            if segment_start <= 30 * 60:
                score += 6
            elif segment_start <= 35 * 60:
                score += 2
            else:
                score -= 8

            window_start = max(0.0, segment_start - PRE_START_CONTEXT_WINDOW_SECONDS)
            context_segments = [
                prev for prev in segments
                if window_start <= float(prev.get("start", 0.0)) < segment_start
            ]
            context_text = " ".join(str(prev.get("text", "")).strip().lower() for prev in context_segments)
            if any(marker in context_text for marker in lowered_context_markers):
                score += 5

            lookahead_segments = segments[index + 1:index + 7]
            lookahead_text = " ".join(
                str(next_segment.get("text", "")).strip().lower()
                for next_segment in lookahead_segments
            )
            if contains_marker(lookahead_text, prayer_markers):
                score += 3
            if contains_marker(lookahead_text, START_SERMON_TEXT_MARKERS):
                score += 3
            if contains_marker(lookahead_text, START_INTRO_MARKERS):
                score += 8
            if has_opening_prayer_content and contains_marker(lookahead_text, OPENING_PRAYER_FOLLOWUP_MARKERS):
                score += 24
                candidate_reason = f"{candidate_reason} + opening-prayer sermon followup"

            post_prayer_window_segments = []
            for next_segment in segments[index + 1:]:
                next_start = float(next_segment.get("start", 0.0))
                if next_start > segment_start + 180:
                    break
                post_prayer_window_segments.append(next_segment)
            post_prayer_window_text = " ".join(
                str(next_segment.get("text", "")).strip().lower()
                for next_segment in post_prayer_window_segments
            )
            has_sermon_followup = contains_marker(post_prayer_window_text, PRE_SERMON_SEQUENCE_MARKERS)
            has_strong_sermon_followup = contains_marker(post_prayer_window_text, STRONG_SERMON_FOLLOWUP_MARKERS)
            has_scripture_reading_followup = contains_marker(post_prayer_window_text, SCRIPTURE_READING_MARKERS)
            prayer_continuation_hits = sum(
                1 for marker in lowered_prayer_continuation_markers if marker in post_prayer_window_text
            )
            if has_prayer_start and has_strong_sermon_followup:
                score += 18
                candidate_reason = f"{candidate_reason} + strong sermon followup"
            elif has_prayer_start and has_sermon_followup:
                score += 4
                candidate_reason = f"{candidate_reason} + weak sermon followup"
            elif has_prayer_start:
                score -= 18
                candidate_reason = f"{candidate_reason} - no sermon followup"
            if has_prayer_start and has_scripture_reading_followup and not has_strong_sermon_followup:
                score -= 12
                candidate_reason = f"{candidate_reason} - scripture reading"
            if has_prayer_start and prayer_continuation_hits >= 2 and not has_strong_sermon_followup:
                score -= 10
                candidate_reason = f"{candidate_reason} - prayer continuation"

            # Sermon opening prayer is usually preceded by a short congregational response like "아멘".
            near_context_segments = segments[max(0, index - 3):index]
            near_context_text = " ".join(
                str(prev_segment.get("text", "")).strip().lower()
                for prev_segment in near_context_segments
            )
            if has_prayer_start and contains_marker(near_context_text, PRE_PRAYER_TRANSITION_MARKERS):
                score += 5
                candidate_reason = f"{candidate_reason} + pre-prayer response"
            if has_bow_prayer_start:
                score += 6

            if has_prayer_start and not has_choir_context:
                self.log(
                    f"Rejected start candidate {format_timestamp(segment_start)}: "
                    "prayer marker before choir context."
                )
                continue

            if has_prayer_start and audio_file and ffmpeg_path and score >= 16:
                start_candidates.append(StartCandidate(score=score, time_seconds=candidate_start, reason=candidate_reason))
            elif has_applause and score >= 16:
                start_candidates.append(StartCandidate(score=score, time_seconds=candidate_start, reason="transcript applause"))
            elif score >= 14:
                start_candidates.append(StartCandidate(score=score, time_seconds=candidate_start, reason=candidate_reason))

        if start_time is None and start_candidates:
            start_candidates.sort(key=lambda item: item.score, reverse=True)
            if source_file and api_key:
                refined_time = self.refine_start_with_vision(
                    source_file=source_file,
                    ffmpeg_path=ffmpeg_path,
                    api_key=api_key,
                    vision_model=vision_model,
                    candidates=start_candidates,
                    title_slug=title_slug,
                )
                if refined_time is not None:
                    start_time = refined_time
            if start_time is None:
                if source_file and api_key:
                    raise ValueError(
                        "Could not confirm the exact FGMC sermon title slide. "
                        "The app intentionally rejected transcript-only start candidates because they can confuse prayer slides like '합심기도', 대표기도, or 성경봉독. "
                        "Check the vision logs or enter the start time manually."
                    )
                start_time = start_candidates[0].time_seconds

        if start_time is None:
            raise ValueError("Could not detect a sermon start marker. Enter times manually.")

        end_candidates: list[tuple[int, float]] = []
        prayer_only_candidates: list[tuple[int, float]] = []
        for index, segment in enumerate(segments):
            text = str(segment.get("text", "")).strip().lower()
            segment_start = float(segment.get("start", 0.0))
            if segment_start < start_time + MIN_SERMON_DURATION_SECONDS:
                continue
            if segment_start > start_time + MAX_SERMON_DURATION_SECONDS:
                break
            normalized_text = normalize_search_text(text)
            is_direct_blessing = any(marker in text for marker in lowered_end_markers)
            is_fuzzy_blessing = any(
                normalize_search_text(marker) in normalized_text for marker in fuzzy_end_markers
            )
            has_blessing_style_phrase = (
                "\uc8fc\ub2d8\uc758\uc774\ub984\uc73c\ub85c" in normalized_text
                and ("\ub418\uc2dc\uae38" in normalized_text or "\ub418\uc2dc\uae30\ub97c" in normalized_text)
            )
            has_closing_prayer_transition = contains_marker(
                text,
                [
                    "\uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
                    "\ucc2c\uc591\ud560\uae4c\uc694",
                    "\uadf8\ub9ac\uace0 \uae30\ub3c4\ud558\uaca0\uc2b5\ub2c8\ub2e4",
                ],
            )
            if not (is_direct_blessing or is_fuzzy_blessing or has_blessing_style_phrase):
                continue

            candidate_score = 10 if is_direct_blessing else 6
            candidate_end = float(segment.get("end", segment.get("start", segment_start)))
            prayer_window_end = candidate_end + POST_END_CONTEXT_WINDOW_SECONDS
            last_matching_end = candidate_end
            found_prayer_close = False

            if has_blessing_style_phrase:
                candidate_score += 4
            if has_closing_prayer_transition:
                candidate_score += 8

            for next_segment in segments[index:]:
                next_start = float(next_segment.get("start", 0.0))
                if next_start > prayer_window_end:
                    break
                next_text = str(next_segment.get("text", "")).strip().lower()
                next_end = float(next_segment.get("end", next_segment.get("start", next_start)))
                if contains_marker(next_text, END_PRAYER_MARKERS):
                    candidate_score += 5
                    last_matching_end = max(last_matching_end, next_end)
                    found_prayer_close = True

            if segment_start >= start_time + (20 * 60):
                candidate_score += 2

            end_candidates.append((candidate_score, last_matching_end))

            if found_prayer_close:
                prayer_only_candidates.append((candidate_score, last_matching_end))

        if end_candidates:
            _, end_time = max(end_candidates, key=lambda item: item[0])

        if end_time is None:
            trailing_prayer_candidates: list[tuple[int, float]] = []
            for segment in segments:
                segment_start = float(segment.get("start", 0.0))
                if segment_start < start_time + MIN_SERMON_DURATION_SECONDS:
                    continue
                if segment_start > start_time + MAX_SERMON_DURATION_SECONDS:
                    break
                text = str(segment.get("text", "")).strip().lower()
                if contains_marker(text, END_PRAYER_MARKERS):
                    score = 5
                    if segment_start >= start_time + (20 * 60):
                        score += 2
                    trailing_prayer_candidates.append(
                        (score, float(segment.get("end", segment.get("start", segment_start))))
                    )

            if prayer_only_candidates:
                _, end_time = max(prayer_only_candidates, key=lambda item: item[0])
                self.log("End marker fallback used: closing prayer end without a clear '異뺤썝' transcript match.")
            elif trailing_prayer_candidates:
                _, end_time = max(trailing_prayer_candidates, key=lambda item: item[0])
                self.log("End marker fallback used: selected the latest closing-prayer style ending.")
            else:
                tail_segments = [
                    segment for segment in segments
                    if start_time + MIN_SERMON_DURATION_SECONDS
                    <= float(segment.get("start", 0.0))
                    <= start_time + MAX_SERMON_DURATION_SECONDS
                ]
                if tail_segments:
                    tail_segment = tail_segments[-1]
                    end_time = float(tail_segment.get("end", tail_segment.get("start", start_time + MIN_SERMON_DURATION_SECONDS)))
                    self.log("End marker fallback used: selected the latest transcript segment in the sermon window.")
                else:
                    final_segment = segments[-1]
                    end_time = float(
                        final_segment.get("end", final_segment.get("start", start_time + MIN_SERMON_DURATION_SECONDS))
                    )
                    self.log("End marker fallback used: selected the final transcript segment because no end candidates matched.")
        return format_timestamp(start_time), format_timestamp(end_time)

    def extract_title_frame(
        self, source_file: Path, title_slug: str, ffmpeg_path: str, start_time_text: str
    ) -> Path:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        capture_seconds = parse_timecode(start_time_text)
        frame_path = FRAME_DIR / f"{title_slug}-title-frame.jpg"
        command = [
            ffmpeg_bin,
            "-y",
            "-ss",
            format_timestamp(capture_seconds),
            "-i",
            str(source_file),
            "-frames:v",
            "1",
            "-q:v",
            "2",
            str(frame_path),
        ]
        self.status("Capturing frame for sermon title suggestion...")
        self.progress(None)
        self.log("Capturing frame for sermon title suggestion...")
        run_command(command)
        return frame_path

    def suggest_title_from_frame(self, frame_path: Path, api_key: str, vision_model: str) -> str:
        if not api_key:
            raise ValueError("OpenAI API key is missing. Add it in Settings.")

        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        image_data = base64.b64encode(frame_path.read_bytes()).decode("ascii")
        response = client.responses.create(
            model=vision_model or "gpt-4.1-mini",
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "Read the sermon title visible on this Korean church sermon title slide. "
                                "The slide may also show a scripture reference and pastor name. "
                                "Return only the sermon title text, not the scripture reference, not the pastor name, "
                                "and no explanation."
                            ),
                        },
                        {
                            "type": "input_image",
                            "image_url": f"data:image/jpeg;base64,{image_data}",
                            "detail": "high",
                        },
                    ],
                }
            ],
        )
        title = getattr(response, "output_text", "").strip()
        if not title:
            raise ValueError("Could not infer a sermon title from the frame.")
        return title

    def get_transcript_path(self, title_slug: str) -> Path:
        return TRANSCRIPT_DIR / f"{title_slug}.json"

    def describe_transcript_tail(self, transcript_path: Path, limit: int = 8) -> list[str]:
        payload = json.loads(transcript_path.read_text(encoding="utf-8"))
        segments = payload.get("segments") or []
        if not segments:
            return ["Transcript has no timestamped segments."]

        lines = [
            f"Transcript model: {payload.get('model', 'unknown')}",
            f"Transcript segments: {len(segments)}",
            "Latest transcript tail:",
        ]
        for segment in segments[-limit:]:
            start = format_timestamp(float(segment.get("start", 0.0)))
            end = format_timestamp(float(segment.get("end", segment.get("start", 0.0))))
            text = str(segment.get("text", "")).strip()
            lines.append(f"{start} - {end} | {text}")
        return lines

    def get_media_duration(self, source_file: Path, ffmpeg_path: str) -> str:
        ffprobe_bin = resolve_ffprobe(resolve_binary("ffmpeg", ffmpeg_path))
        command = [
            ffprobe_bin,
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(source_file),
        ]
        completed = run_command(command)
        return format_timestamp(float(completed.stdout.strip()))

    def export_clip(
        self,
        source_file: Path,
        title: str,
        start_text: str,
        end_text: str,
        ffmpeg_path: str,
        fade_out_enabled: bool = True,
        fade_out_seconds_text: str = "3",
    ) -> Path:
        start_seconds = parse_timecode(start_text)
        end_seconds = parse_timecode(end_text)
        if end_seconds <= start_seconds:
            raise ValueError("End time must be after start time.")

        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        destination = EXPORTS_DIR / f"{slugify(title or 'sermon')}.mp4"
        duration_seconds = end_seconds - start_seconds
        fade_out_seconds = 0.0
        if fade_out_enabled:
            fade_out_seconds = float(fade_out_seconds_text.strip() or "3")
            if fade_out_seconds < 0:
                raise ValueError("Fade-out seconds must be zero or a positive number.")
            fade_out_seconds = min(fade_out_seconds, max(0.0, duration_seconds - 0.5))

        command = [
            ffmpeg_bin,
            "-y",
            "-i",
            str(source_file),
            "-ss",
            format_timestamp(start_seconds),
            "-to",
            format_timestamp(end_seconds),
        ]
        if fade_out_seconds > 0:
            fade_start = max(0.0, duration_seconds - fade_out_seconds)
            command.extend(
                [
                    "-vf",
                    f"setpts=PTS-STARTPTS,fade=t=out:st={fade_start:.3f}:d={fade_out_seconds:.3f}",
                    "-af",
                    f"asetpts=PTS-STARTPTS,afade=t=out:st={fade_start:.3f}:d={fade_out_seconds:.3f}",
                ]
            )
        command.extend(
            [
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
                "-progress",
                "pipe:1",
                "-nostats",
                str(destination),
            ]
        )
        self.status("Exporting sermon MP4...")
        self.progress(0, 100)
        if fade_out_seconds > 0:
            self.log(f"Exporting sermon MP4 with {fade_out_seconds:g}s fade-out...")
        else:
            self.log("Exporting sermon MP4...")

        def handle_export_progress(line: str) -> None:
            progress_seconds = parse_ffmpeg_progress_seconds(line)
            if progress_seconds is None:
                return
            percent = min(100, max(0, int((progress_seconds / duration_seconds) * 100)))
            self.status(f"Exporting sermon MP4... {percent}%")
            self.progress(percent, 100)

        run_command_streaming(command, handle_export_progress)
        self.progress(100, 100)
        return destination

    def resize_mp4(
        self,
        source_file: Path,
        target_mb_text: str,
        ffmpeg_path: str,
    ) -> Path:
        target_mb = float(target_mb_text.strip())
        if target_mb <= 0:
            raise ValueError("Target file size must be a positive number.")

        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        destination = source_file.with_name(f"{source_file.stem}-{int(target_mb)}mb.mp4")
        try:
            duration_seconds = parse_timecode(self.get_media_duration(source_file, ffmpeg_path))
        except Exception:
            duration_seconds = 0.0
        if duration_seconds <= 0:
            raise ValueError("Could not read the MP4 duration for target-size conversion.")

        audio_kbps = 128
        total_kbps = int((target_mb * 1024 * 8) / duration_seconds)
        video_kbps = max(300, total_kbps - audio_kbps)
        if video_kbps <= 300 and total_kbps <= audio_kbps + 300:
            self.log(
                f"Target size {target_mb:g}MB is very small for {format_seconds(duration_seconds)}; "
                "using minimum video bitrate."
            )

        command = [
            ffmpeg_bin,
            "-y",
            "-i",
            str(source_file),
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-b:v",
            f"{video_kbps}k",
            "-c:a",
            "aac",
            "-b:a",
            f"{audio_kbps}k",
            "-movflags",
            "+faststart",
            "-progress",
            "pipe:1",
            "-nostats",
            str(destination),
        ]
        self.status(f"Reducing MP4 near {target_mb:g}MB...")
        self.progress(0, 100)
        self.log(
            f"Reducing MP4 near {target_mb:g}MB: {source_file} "
            f"(video {video_kbps}k, audio {audio_kbps}k)"
        )

        def handle_resize_progress(line: str) -> None:
            progress_seconds = parse_ffmpeg_progress_seconds(line)
            if progress_seconds is None or duration_seconds <= 0:
                return
            percent = min(100, max(0, int((progress_seconds / duration_seconds) * 100)))
            self.status(f"Reducing MP4 near {target_mb:g}MB... {percent}%")
            self.progress(percent, 100)

        run_command_streaming(command, handle_resize_progress)
        self.progress(100, 100)
        return destination

    def create_review_clip(
        self,
        source_file: Path,
        ffmpeg_path: str,
        center_time_text: str,
        title_slug: str,
        label: str,
        seconds_before: int = 5,
        seconds_after: int = 5,
    ) -> Path:
        ffmpeg_bin = resolve_binary("ffmpeg", ffmpeg_path)
        center_seconds = parse_timecode(center_time_text)
        clip_start = max(0.0, center_seconds - seconds_before)
        clip_duration = max(2, seconds_before + seconds_after)
        destination = REVIEW_DIR / f"{title_slug}-{label}-review.mp4"
        command = [
            ffmpeg_bin,
            "-y",
            "-ss",
            format_timestamp(clip_start),
            "-i",
            str(source_file),
            "-t",
            str(clip_duration),
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "22",
            "-c:a",
            "aac",
            "-b:a",
            "128k",
            "-movflags",
            "+faststart",
            str(destination),
        ]
        self.status(f"Creating {label} review clip...")
        self.progress(None)
        self.log(f"Creating {label} review clip...")
        run_command(command)
        return destination


class WorkerThread(QThread):
    log = Signal(str)
    status = Signal(str)
    progress = Signal(object, object)
    finished_ok = Signal(object)
    failed = Signal(str, str)

    def __init__(self, fn, *args, error_title="Error", **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.error_title = error_title

    def run(self) -> None:
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.finished_ok.emit(result)
        except Exception as exc:
            self.failed.emit(self.error_title, str(exc))


class SettingsDialog(QDialog):
    def __init__(self, engine: SermonStudioEngine, parent=None):
        super().__init__(parent)
        self.engine = engine
        self.setWindowTitle("설정")
        self.setModal(True)

        self.yt_dlp_edit = QLineEdit(engine.get_setting("yt_dlp_path"))
        self.ffmpeg_edit = QLineEdit(engine.get_setting("ffmpeg_path"))
        self.youtube_client_secrets_edit = QLineEdit(engine.get_setting("youtube_client_secrets_path"))
        self.api_key_edit = QLineEdit()
        if engine.get_setting("openai_api_key"):
            self.api_key_edit.setPlaceholderText("OPENAI_API_KEY is already saved in Windows")
        else:
            self.api_key_edit.setPlaceholderText("Enter key once to save to Windows user environment")
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.transcription_model_edit = QLineEdit(
            engine.get_setting("transcription_model", "whisper-1")
        )
        self.language_edit = QLineEdit(engine.get_setting("language", "ko"))
        self.vision_model_edit = QLineEdit(engine.get_setting("vision_model", "gpt-4.1-mini"))

        layout = QVBoxLayout()
        form = QFormLayout()
        form.addRow("yt-dlp path", self._with_browse(self.yt_dlp_edit))
        form.addRow("ffmpeg path", self._with_browse(self.ffmpeg_edit))
        form.addRow("YouTube OAuth JSON", self._with_browse_json(self.youtube_client_secrets_edit))
        form.addRow("OpenAI API key", self.api_key_edit)
        form.addRow("Transcription model", self.transcription_model_edit)
        form.addRow("Language code", self.language_edit)
        form.addRow("Vision model", self.vision_model_edit)
        layout.addLayout(form)

        note = QLabel(
            "ffmpeg.exe 와 yt-dlp.exe 가 이 프로그램과 같은 폴더 또는 옆의 tools 폴더에 있으면 자동으로 찾습니다.\n"
            "OpenAI API key 는 settings.json 에 저장하지 않고 Windows 사용자 환경변수 OPENAI_API_KEY 로 저장합니다.\n"
            "YouTube OAuth JSON은 Google Cloud에서 받은 데스크톱 앱용 client secrets 파일입니다."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #555;")
        layout.addWidget(note)

        row = QHBoxLayout()
        save_btn = QPushButton("저장")
        close_btn = QPushButton("닫기")
        save_btn.clicked.connect(self.save)
        close_btn.clicked.connect(self.reject)
        row.addStretch(1)
        row.addWidget(save_btn)
        row.addWidget(close_btn)
        layout.addLayout(row)
        self.setLayout(layout)

    def _with_browse(self, line_edit: QLineEdit) -> QWidget:
        wrapper = QWidget()
        row = QHBoxLayout(wrapper)
        row.setContentsMargins(0, 0, 0, 0)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse(line_edit))
        row.addWidget(line_edit)
        row.addWidget(browse)
        return wrapper

    def _with_browse_json(self, line_edit: QLineEdit) -> QWidget:
        wrapper = QWidget()
        row = QHBoxLayout(wrapper)
        row.setContentsMargins(0, 0, 0, 0)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse_json(line_edit))
        row.addWidget(line_edit)
        row.addWidget(browse)
        return wrapper

    def _browse(self, line_edit: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "실행 파일 선택", "", "Executable (*.exe)")
        if selected:
            line_edit.setText(selected)

    def _browse_json(self, line_edit: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Google OAuth JSON 선택",
            "",
            "JSON files (*.json);;All files (*.*)",
        )
        if selected:
            line_edit.setText(selected)

    def save(self) -> None:
        api_key = self.api_key_edit.text().strip()
        key_saved = False
        try:
            if api_key:
                self.engine.save_openai_api_key(api_key)
                key_saved = True
        except Exception as exc:
            QMessageBox.critical(self, "API key 저장 실패", f"OpenAI API key를 Windows 환경변수에 저장하지 못했습니다.\n\n{exc}")
            return
        self.engine.save_settings(
            {
                "yt_dlp_path": self.yt_dlp_edit.text().strip(),
                "ffmpeg_path": self.ffmpeg_edit.text().strip(),
                "youtube_client_secrets_path": self.youtube_client_secrets_edit.text().strip(),
                "transcription_model": self.transcription_model_edit.text().strip() or "whisper-1",
                "language": self.language_edit.text().strip() or "ko",
                "vision_model": self.vision_model_edit.text().strip() or "gpt-4.1-mini",
            }
        )
        if key_saved:
            QMessageBox.information(self, "저장 완료", "OpenAI API key가 Windows 사용자 환경변수 OPENAI_API_KEY에 저장되었습니다.\n프로그램을 재실행하면 가장 안전합니다.")
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FGMC (체리힐제일교회) YOUTUBE AUTOMATION PROGRAM")
        self.resize(980, 760)
        self.engine = SermonStudioEngine(self.log)
        self.current_source_file: Path | None = None
        self.last_export_file: Path | None = None
        self.current_title_slug = ""
        self.active_thread: WorkerThread | None = None

        self.url_edit = QLineEdit()
        self.title_edit = QLineEdit()
        self.speaker_edit = QLineEdit()
        self.date_edit = QLineEdit()
        self.chunk_minutes_edit = QLineEdit("10")
        self.duration_edit = QLineEdit("-")
        self.duration_edit.setReadOnly(True)
        self.start_edit = QLineEdit()
        self.end_edit = QLineEdit()
        self.review_seconds_edit = QLineEdit("5")
        self.fade_out_check = QCheckBox("끝부분 자연스럽게 마무리")
        self.fade_out_check.setChecked(True)
        self.fade_out_seconds_edit = QLineEdit("3")
        self.target_size_mb_edit = QLineEdit("100")
        self.status_label = QLabel("Ready.")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.manual_actions_widget: QWidget | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        main = QVBoxLayout(root)

        top_row = QHBoxLayout()
        title = QLabel("FGMC (체리힐제일교회) YOUTUBE AUTOMATION PROGRAM")
        title.setStyleSheet("font-size: 22px; font-weight: 700;")
        settings_btn = QPushButton("설정")
        settings_btn.clicked.connect(self.open_settings)
        top_row.addWidget(title)
        top_row.addStretch(1)
        top_row.addWidget(settings_btn)
        main.addLayout(top_row)

        info_box = QGroupBox("예배 영상")
        info_form = QFormLayout()
        info_form.addRow("유튜브 URL", self.url_edit)
        info_form.addRow("설교 제목", self.title_edit)
        info_form.addRow("설교자", self.speaker_edit)
        info_form.addRow("날짜", self.date_edit)
        info_box.setLayout(info_form)
        main.addWidget(info_box)

        live_row = QHBoxLayout()
        find_first_live_btn = QPushButton("1부 예배 찾기")
        find_second_live_btn = QPushButton("2부 예배 찾기")
        find_first_live_btn.clicked.connect(lambda: self.find_live_archive("first"))
        find_second_live_btn.clicked.connect(lambda: self.find_live_archive("second"))
        live_row.addStretch(1)
        live_row.addWidget(find_first_live_btn)
        live_row.addWidget(find_second_live_btn)
        live_row.addStretch(1)
        main.addLayout(live_row)

        mode_row = QHBoxLayout()
        auto_btn = QPushButton("자동 업로드")
        manual_btn = QPushButton("수동 업로드")
        auto_btn.setMinimumHeight(44)
        auto_btn.setMinimumWidth(180)
        manual_btn.setMinimumHeight(44)
        manual_btn.setMinimumWidth(180)
        auto_btn.setStyleSheet(
            "QPushButton {"
            "background-color: #1f6f43; color: white; border: 0; border-radius: 8px;"
            "font-size: 16px; font-weight: 700; padding: 10px 24px;"
            "}"
            "QPushButton:hover { background-color: #258250; }"
            "QPushButton:pressed { background-color: #185735; }"
        )
        manual_btn.setStyleSheet(
            "QPushButton {"
            "background-color: #f2efe8; color: #29352b; border: 1px solid #c8c0b3; border-radius: 8px;"
            "font-size: 16px; font-weight: 700; padding: 10px 24px;"
            "}"
            "QPushButton:hover { background-color: #e7dfd2; }"
            "QPushButton:pressed { background-color: #d8cebf; }"
        )
        auto_btn.clicked.connect(self.auto_upload)
        manual_btn.clicked.connect(self.toggle_manual_upload)
        mode_row.addStretch(1)
        mode_row.addWidget(auto_btn)
        mode_row.addWidget(manual_btn)
        mode_row.addStretch(1)
        main.addLayout(mode_row)

        self.manual_actions_widget = QWidget()
        action_row = QHBoxLayout()
        download_btn = QPushButton("1. Download Full Video")
        choose_local_btn = QPushButton("로컬 영상 선택")
        detect_btn = QPushButton("2. Auto Detect Sermon Range")
        detect_skip_btn = QPushButton("2b. Reuse Transcript Only")
        export_btn = QPushButton("3. Export Sermon MP4")
        open_full_btn = QPushButton("풀영상 보기")
        download_btn.clicked.connect(self.download_video)
        choose_local_btn.clicked.connect(self.choose_local_video)
        detect_btn.clicked.connect(self.auto_detect)
        detect_skip_btn.clicked.connect(self.auto_detect_skip_transcription)
        export_btn.clicked.connect(self.export_clip)
        open_full_btn.clicked.connect(self.open_full_video)
        action_row.addWidget(download_btn)
        action_row.addWidget(choose_local_btn)
        action_row.addWidget(open_full_btn)
        action_row.addWidget(detect_btn)
        action_row.addWidget(detect_skip_btn)
        action_row.addWidget(export_btn)
        action_row.addStretch(1)
        self.manual_actions_widget.setLayout(action_row)
        self.manual_actions_widget.setVisible(False)
        main.addWidget(self.manual_actions_widget)

        clip_box = QGroupBox("설교 구간")
        clip_grid = QGridLayout()
        clip_grid.addWidget(QLabel("청크 분"), 0, 0)
        clip_grid.addWidget(self.chunk_minutes_edit, 0, 1)
        clip_grid.addWidget(QLabel("전체 길이"), 0, 2)
        clip_grid.addWidget(self.duration_edit, 0, 3)
        clip_grid.addWidget(QLabel("시작 시간"), 1, 0)
        clip_grid.addWidget(self.start_edit, 1, 1)
        clip_grid.addWidget(QLabel("종료 시간"), 1, 2)
        clip_grid.addWidget(self.end_edit, 1, 3)
        clip_grid.addWidget(QLabel("리뷰 초"), 2, 0)
        clip_grid.addWidget(self.review_seconds_edit, 2, 1)
        clip_grid.addWidget(QLabel("목표 파일 크기(MB)"), 2, 2)
        clip_grid.addWidget(self.target_size_mb_edit, 2, 3)
        clip_grid.addWidget(self.fade_out_check, 3, 0)
        clip_grid.addWidget(QLabel("페이드아웃 초"), 3, 1)
        clip_grid.addWidget(self.fade_out_seconds_edit, 3, 2)
        resize_btn = QPushButton("목표 크기로 줄이기")
        resize_btn.clicked.connect(self.resize_exported_mp4)
        clip_grid.addWidget(resize_btn, 3, 3)

        start_controls = self._adjust_controls(self.start_edit, self.review_start, "Review Start")
        end_controls = self._adjust_controls(self.end_edit, self.review_end, "Review End")
        clip_grid.addLayout(start_controls, 4, 0, 1, 2)
        clip_grid.addLayout(end_controls, 4, 2, 1, 2)
        clip_box.setLayout(clip_grid)
        main.addWidget(clip_box)

        progress_box = QGroupBox("진행 상태")
        progress_layout = QVBoxLayout()
        progress_layout.addWidget(self.status_label)
        progress_layout.addWidget(self.progress_bar)
        progress_box.setLayout(progress_layout)
        main.addWidget(progress_box)

        log_box = QGroupBox("로그")
        log_layout = QVBoxLayout()
        log_layout.addWidget(self.log_box)
        log_box.setLayout(log_layout)
        main.addWidget(log_box, 1)

        self.setCentralWidget(root)

    def _adjust_controls(self, line_edit: QLineEdit, review_handler, review_label: str) -> QHBoxLayout:
        row = QHBoxLayout()
        for label, delta in [("-10s", -10), ("-5s", -5), ("+5s", 5), ("+10s", 10)]:
            button = QPushButton(label)
            button.clicked.connect(lambda _=False, target=line_edit, d=delta: self.shift_time(target, d))
            row.addWidget(button)
        review_button = QPushButton(review_label)
        review_button.clicked.connect(review_handler)
        row.addWidget(review_button)
        row.addStretch(1)
        return row

    def log(self, message: str) -> None:
        self.log_box.append(message)
        cursor = self.log_box.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_box.setTextCursor(cursor)

    def set_status(self, message: str) -> None:
        self.status_label.setText(message or "Working...")

    def set_progress(self, value: int | None, maximum: int | None = None) -> None:
        if value is None:
            self.progress_bar.setRange(0, 0)
            return
        self.progress_bar.setRange(0, maximum or 100)
        self.progress_bar.setValue(value)

    def task_log(self, message: str) -> None:
        if self.active_thread:
            self.active_thread.log.emit(message)
        else:
            self.log(message)

    def task_status(self, message: str) -> None:
        if self.active_thread:
            self.active_thread.status.emit(message)
        else:
            self.set_status(message)

    def task_progress(self, value: int | None, maximum: int | None = None) -> None:
        if self.active_thread:
            self.active_thread.progress.emit(value, maximum)
        else:
            self.set_progress(value, maximum)

    def open_settings(self) -> None:
        dialog = SettingsDialog(self.engine, self)
        if dialog.exec():
            self.log("Settings saved.")

    def toggle_manual_upload(self) -> None:
        if self.manual_actions_widget:
            self.manual_actions_widget.setVisible(not self.manual_actions_widget.isVisible())

    def show_error(self, title: str, body: str) -> None:
        QMessageBox.critical(self, title, body)

    def show_info(self, title: str, body: str) -> None:
        QMessageBox.information(self, title, body)

    def start_worker(self, fn, *args, error_title="Error", on_success=None) -> None:
        if self.active_thread and self.active_thread.isRunning():
            self.show_error("Busy", "Another task is still running.")
            return
        self.engine.set_callbacks(self.task_log, self.task_status, self.task_progress)
        worker = WorkerThread(fn, *args, error_title=error_title)
        self.active_thread = worker
        self.set_status("Starting task...")
        self.set_progress(None)
        worker.log.connect(self.log)
        worker.status.connect(self.set_status)
        worker.progress.connect(self.set_progress)
        worker.failed.connect(self.show_error)
        worker.failed.connect(lambda *_: self._finish_task("Failed."))
        if on_success:
            worker.finished_ok.connect(on_success)
        worker.finished_ok.connect(lambda *_: self._finish_task("Done."))
        worker.start()

    def _finish_task(self, status_message: str) -> None:
        self.set_status(status_message)
        self.set_progress(100, 100)
        self.active_thread = None

    def shift_time(self, widget: QLineEdit, delta_seconds: int) -> None:
        raw = widget.text().strip()
        if not raw:
            return
        try:
            widget.setText(format_timestamp(max(0.0, parse_timecode(raw) + delta_seconds)))
        except ValueError:
            self.show_error("Invalid timecode", f"Could not parse timecode: {raw}")

    def find_live_archive(self, service_order: str) -> None:
        self.start_worker(
            lambda: self.engine.find_recent_live_archive(service_order),
            error_title="라이브 아카이브 찾기 실패",
            on_success=self._after_find_live_archive,
        )

    def _after_find_live_archive(self, candidate: YouTubeLiveCandidate) -> None:
        self.url_edit.setText(candidate.url)
        self.log(f"Live archive selected: {candidate.url}")
        message = f"URL 칸에 라이브 아카이브를 입력했습니다:\n{candidate.url}"
        if candidate.title:
            message += f"\n\n라이브 제목: {candidate.title}"
        if candidate.processing_status and candidate.processing_status != "unknown":
            message += f"\n\nYouTube processing status: {candidate.processing_status}"
        self.show_info("라이브 아카이브 찾기 완료", message)

    def download_video(self) -> None:
        self.start_worker(
            self._do_download,
            error_title="Download failed",
            on_success=self._after_download,
        )

    def auto_upload(self) -> None:
        self.start_worker(
            self._do_auto_upload,
            error_title="자동 업로드 실패",
            on_success=self._after_auto_upload,
        )

    def _do_auto_upload(self):
        result = self.engine.download_video(
            url=self.url_edit.text().strip(),
            title=self.title_edit.text().strip(),
            yt_dlp_path=self.engine.get_setting("yt_dlp_path"),
        )
        duration = self.engine.get_media_duration(result.source_file, self.engine.get_setting("ffmpeg_path"))
        title_slug = result.title_slug
        full_audio = self.engine.extract_full_audio(
            source_file=result.source_file,
            title_slug=title_slug,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
        )
        chunks = self.engine.split_audio_chunks(
            audio_file=full_audio,
            title_slug=title_slug,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            chunk_minutes=int(self.chunk_minutes_edit.text().strip() or "10"),
        )
        transcript_path = self.engine.transcribe_chunks(
            transcript_name=title_slug,
            model_name=self.engine.get_setting("transcription_model", "whisper-1"),
            api_key=self.engine.get_setting("openai_api_key"),
            language=self.engine.get_setting("language", "ko"),
            chunks=chunks,
        )
        start_text, end_text = self.engine.suggest_sermon_range(
            transcript_path,
            audio_file=full_audio,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            source_file=result.source_file,
            api_key=self.engine.get_setting("openai_api_key"),
            vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            title_slug=title_slug,
        )
        suggested_title = self.title_edit.text().strip()
        try:
            frame_path = self.engine.extract_title_frame(
                source_file=result.source_file,
                title_slug=title_slug,
                ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
                start_time_text=start_text,
            )
            detected_title = self.engine.suggest_title_from_frame(
                frame_path=frame_path,
                api_key=self.engine.get_setting("openai_api_key"),
                vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            )
            if detected_title:
                suggested_title = detected_title
        except Exception as exc:
            self.task_log(f"Title suggestion skipped: {exc}")
        export_title = suggested_title or self.title_edit.text().strip() or title_slug
        destination = self.engine.export_clip(
            source_file=result.source_file,
            title=export_title,
            start_text=start_text,
            end_text=end_text,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            fade_out_enabled=self.fade_out_check.isChecked(),
            fade_out_seconds_text=self.fade_out_seconds_edit.text().strip() or "3",
        )
        return result, duration, transcript_path, start_text, end_text, suggested_title, destination

    def _after_auto_upload(self, payload) -> None:
        result, duration, transcript_path, start_text, end_text, suggested_title, destination = payload
        self.current_source_file = result.source_file
        self.current_title_slug = result.title_slug
        self.duration_edit.setText(duration)
        self.start_edit.setText(start_text)
        self.end_edit.setText(end_text)
        if suggested_title:
            self.title_edit.setText(suggested_title)
        self.log(f"Downloaded file: {result.source_file}")
        self.log(f"Transcript saved: {transcript_path}")
        self.log(f"Exported file: {destination}")
        self.last_export_file = destination
        self.show_info("Done", f"MP4 생성 완료:\n{destination}")

    def _do_download(self):
        result = self.engine.download_video(
            url=self.url_edit.text().strip(),
            title=self.title_edit.text().strip(),
            yt_dlp_path=self.engine.get_setting("yt_dlp_path"),
        )
        duration = self.engine.get_media_duration(result.source_file, self.engine.get_setting("ffmpeg_path"))
        return result, duration

    def _after_download(self, payload) -> None:
        result, duration = payload
        self.current_source_file = result.source_file
        self.current_title_slug = result.title_slug
        self.duration_edit.setText(duration)
        self.log(f"Downloaded file: {result.source_file}")
        self.show_info("Done", "Full service video download finished.")

    def choose_local_video(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "풀영상 파일 선택",
            "",
            "Video files (*.mp4 *.mkv *.mov *.webm *.m4v);;All files (*.*)",
        )
        if not selected:
            return
        source_file = Path(selected)
        if not source_file.exists():
            self.show_error("파일 없음", f"선택한 파일을 찾을 수 없습니다:\n{source_file}")
            return
        self.current_source_file = source_file
        self.current_title_slug = slugify(self.title_edit.text().strip() or source_file.stem)
        try:
            duration = self.engine.get_media_duration(source_file, self.engine.get_setting("ffmpeg_path"))
            self.duration_edit.setText(duration)
        except Exception as exc:
            self.duration_edit.setText("-")
            self.log(f"Could not read local video duration: {exc}")
        self.log(f"Selected local full video: {source_file}")
        self.show_info("Done", "로컬 풀영상이 선택되었습니다. 이제 시간을 직접 입력하거나 Auto Detect를 실행할 수 있습니다.")

    def open_full_video(self) -> None:
        source_file = self.current_source_file
        if not source_file:
            title_slug = self.current_title_slug or build_job_slug(
                self.url_edit.text().strip(),
                self.title_edit.text().strip(),
            )
            candidates = sorted(
                path for path in (DOWNLOADS_DIR / title_slug).glob("source.*")
                if not path.name.endswith((".part", ".ytdl", ".temp"))
            )
            source_file = candidates[0] if candidates else None
        if not source_file or not source_file.exists():
            self.show_error("풀영상 없음", "먼저 1. Download Full Video를 실행해 주세요.")
            return
        os.startfile(str(source_file))

    def auto_detect(self) -> None:
        self.start_worker(
            self._do_auto_detect,
            error_title="Auto detect failed",
            on_success=self._after_auto_detect,
        )

    def auto_detect_skip_transcription(self) -> None:
        self.start_worker(
            self._do_auto_detect_skip_transcription,
            error_title="Auto detect failed",
            on_success=self._after_auto_detect,
        )

    def _do_auto_detect(self):
        if not self.current_source_file:
            raise ValueError("Download the full video first.")
        title_slug = self.current_title_slug or build_job_slug(
            self.url_edit.text().strip(),
            self.title_edit.text().strip(),
        )
        full_audio = self.engine.extract_full_audio(
            source_file=self.current_source_file,
            title_slug=title_slug,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
        )
        chunks = self.engine.split_audio_chunks(
            audio_file=full_audio,
            title_slug=title_slug,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            chunk_minutes=int(self.chunk_minutes_edit.text().strip() or "10"),
        )
        transcript_path = self.engine.transcribe_chunks(
            transcript_name=title_slug,
            model_name=self.engine.get_setting("transcription_model", "whisper-1"),
            api_key=self.engine.get_setting("openai_api_key"),
            language=self.engine.get_setting("language", "ko"),
            chunks=chunks,
        )
        start_text, end_text = self.engine.suggest_sermon_range(
            transcript_path,
            audio_file=full_audio,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            source_file=self.current_source_file,
            api_key=self.engine.get_setting("openai_api_key"),
            vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            title_slug=title_slug,
        )
        suggested_title = None
        try:
            frame_path = self.engine.extract_title_frame(
                source_file=self.current_source_file,
                title_slug=title_slug,
                ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
                start_time_text=start_text,
            )
            suggested_title = self.engine.suggest_title_from_frame(
                frame_path=frame_path,
                api_key=self.engine.get_setting("openai_api_key"),
                vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            )
        except Exception as exc:
            self.log(f"Title suggestion skipped: {exc}")
        return transcript_path, start_text, end_text, suggested_title

    def _do_auto_detect_skip_transcription(self):
        if not self.current_source_file:
            raise ValueError("Download the full video first.")
        title_slug = self.current_title_slug or build_job_slug(
            self.url_edit.text().strip(),
            self.title_edit.text().strip(),
        )
        transcript_path = self.engine.get_transcript_path(title_slug)
        if not transcript_path.exists():
            raise FileNotFoundError(
                "No saved transcript was found for this sermon title. Run the full auto detect once first."
            )

        self.task_status("Reusing existing transcript and rerunning sermon detection...")
        self.task_progress(None)
        self.task_log(f"Reusing transcript: {transcript_path}")
        for line in self.engine.describe_transcript_tail(transcript_path):
            self.task_log(line)

        full_audio = self.engine.extract_full_audio(
            source_file=self.current_source_file,
            title_slug=title_slug,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
        )
        start_text, end_text = self.engine.suggest_sermon_range(
            transcript_path,
            audio_file=full_audio,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            source_file=self.current_source_file,
            api_key=self.engine.get_setting("openai_api_key"),
            vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            title_slug=title_slug,
        )
        suggested_title = None
        try:
            frame_path = self.engine.extract_title_frame(
                source_file=self.current_source_file,
                title_slug=title_slug,
                ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
                start_time_text=start_text,
            )
            suggested_title = self.engine.suggest_title_from_frame(
                frame_path=frame_path,
                api_key=self.engine.get_setting("openai_api_key"),
                vision_model=self.engine.get_setting("vision_model", "gpt-4.1-mini"),
            )
        except Exception as exc:
            self.log(f"Title suggestion skipped: {exc}")
        return transcript_path, start_text, end_text, suggested_title

    def _after_auto_detect(self, payload) -> None:
        transcript_path, start_text, end_text, suggested_title = payload
        self.start_edit.setText(start_text)
        self.end_edit.setText(end_text)
        if suggested_title:
            self.title_edit.setText(suggested_title)
            self.log(f"Suggested title: {suggested_title}")
        self.log(f"Transcript saved: {transcript_path}")
        self.show_info("Done", "Suggested sermon timecodes are ready.")

    def review_start(self) -> None:
        self._start_review("start")

    def review_end(self) -> None:
        self._start_review("end")

    def _start_review(self, label: str) -> None:
        self.start_worker(
            lambda: self._do_review(label),
            error_title=f"{label.title()} review failed",
            on_success=self._after_review,
        )

    def _do_review(self, label: str):
        if not self.current_source_file:
            raise ValueError("Download the full video first.")
        time_text = self.start_edit.text().strip() if label == "start" else self.end_edit.text().strip()
        if not time_text:
            raise ValueError(f"Set the {label} time first.")
        title_slug = self.current_title_slug or build_job_slug(
            self.url_edit.text().strip(),
            self.title_edit.text().strip(),
        )
        review_seconds = int(self.review_seconds_edit.text().strip() or "5")
        clip_path = self.engine.create_review_clip(
            source_file=self.current_source_file,
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            center_time_text=time_text,
            title_slug=title_slug,
            label=label,
            seconds_before=review_seconds,
            seconds_after=review_seconds,
        )
        return label, clip_path

    def _after_review(self, payload) -> None:
        label, clip_path = payload
        self.log(f"Review clip ready: {clip_path}")
        os.startfile(str(clip_path))
        self.show_info("Done", f"{label.title()} review clip created.")

    def export_clip(self) -> None:
        if not self.start_edit.text().strip() or not self.end_edit.text().strip():
            self.show_error("Missing timecodes", "Enter start and end time or run auto detect first.")
            return
        self.start_worker(
            self._do_export,
            error_title="Export failed",
            on_success=self._after_export,
        )

    def _do_export(self):
        if not self.current_source_file:
            raise ValueError("Download the full video first.")
        return self.engine.export_clip(
            source_file=self.current_source_file,
            title=self.title_edit.text().strip(),
            start_text=self.start_edit.text().strip(),
            end_text=self.end_edit.text().strip(),
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
            fade_out_enabled=self.fade_out_check.isChecked(),
            fade_out_seconds_text=self.fade_out_seconds_edit.text().strip() or "3",
        )

    def _after_export(self, destination: Path) -> None:
        self.last_export_file = destination
        self.log(f"Exported file: {destination}")
        self.show_info("Done", f"Sermon MP4 saved to:\n{destination}")

    def resize_exported_mp4(self) -> None:
        source_file = self.last_export_file
        if not source_file or not source_file.exists():
            selected, _ = QFileDialog.getOpenFileName(
                self,
                "리사이즈할 MP4 선택",
                str(EXPORTS_DIR),
                "MP4 files (*.mp4);;All files (*.*)",
            )
            if not selected:
                return
            source_file = Path(selected)
        self.start_worker(
            self._do_resize_exported_mp4,
            source_file,
            error_title="MP4 사이즈 변경 실패",
            on_success=self._after_resize_exported_mp4,
        )

    def _do_resize_exported_mp4(self, source_file: Path):
        return self.engine.resize_mp4(
            source_file=source_file,
            target_mb_text=self.target_size_mb_edit.text().strip() or "100",
            ffmpeg_path=self.engine.get_setting("ffmpeg_path"),
        )

    def _after_resize_exported_mp4(self, destination: Path) -> None:
        self.last_export_file = destination
        self.log(f"Resized MP4 saved: {destination}")
        self.show_info("Done", f"사이즈 변경 MP4 저장 완료:\n{destination}")


def main() -> int:
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 10))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
