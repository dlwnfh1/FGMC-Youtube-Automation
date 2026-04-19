# FGMC YouTube Automation Program

Windows desktop program for turning full-length church worship videos into sermon-only MP4 files.

## What It Does

- Downloads a full YouTube worship archive video.
- Extracts audio and creates transcripts with OpenAI.
- Finds sermon start candidates from the transcript.
- Keeps only candidates after choir/praise/special music context.
- Verifies the start point by checking nearby video frames for the sermon title slide.
- Excludes representative prayer, joint prayer, scripture reading, and choir frames.
- Detects the sermon end from transcript phrases such as blessing/closing-prayer markers.
- Exports the final sermon-only MP4 from the original full video.

## Stable Start Detection Logic

Do not change this lightly.

The currently stable and fast start-detection flow is:

1. Reuse or create transcript segments.
2. Generate possible sermon-start candidates from transcript timing.
3. Reject prayer candidates before choir/praise/special-music context.
4. Check only nearby frames around the remaining candidates.
5. Confirm start only when the frame contains a sermon title slide with:
   - sermon title
   - Bible reference
   - preacher role such as pastor or missionary
   - often the heading "말씀"
6. Immediately reject frames showing "대표기도", "합심기도", or "성경봉독".

Avoid returning to full-video 10-second scene scans or contact-sheet start detection unless there is a strong reason.

## Desktop App

Main source files:

- `desktop_app/sermon_studio.py`
- `desktop_app/build_windows.ps1`
- `desktop_app/requirements.txt`

Build output is intentionally not committed. The packaged app should be built locally.

## Local Build

Install Python dependencies, then run:

```powershell
desktop_app\build_windows.ps1
```

The program expects bundled or discoverable:

- `ffmpeg.exe`
- `yt-dlp.exe`

OpenAI API keys are not saved in `settings.json`; they are stored in the Windows user environment variable `OPENAI_API_KEY`.

## Runtime Data

The following are intentionally ignored and should not be committed:

- downloaded videos
- exported MP4 files
- audio chunks
- transcripts
- frame debug images
- built executables
- local tool runtimes

