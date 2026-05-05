from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional

import requests

LOGGER = logging.getLogger(__name__)

import discord
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

opus_path = os.path.join(base_dir, "opus.dll")

if not discord.opus.is_loaded():
    discord.opus.load_opus(opus_path)

print("Opus loaded:", discord.opus.is_loaded())

def _resolve_base_dir(base_dir: Optional[str] = None) -> Path:
    candidates: list[Path] = []

    if base_dir:
        candidates.append(Path(base_dir))

    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent)

    current = Path(__file__).resolve().parent
    candidates.extend(
        [
            current,
            current.parent,
            current.parent.parent,
        ]
    )

    seen: set[Path] = set()
    for cand in candidates:
        try:
            cand = cand.resolve()
        except Exception:
            continue

        if cand in seen:
            continue
        seen.add(cand)

        if (cand / "VOICEVOX" / "run.exe").exists():
            return cand
        if (cand / "FFmpeg" / "ffmpeg.exe").exists():
            return cand
        if (cand / "speakers.json").exists():
            return cand

    if base_dir:
        return Path(base_dir).resolve()

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def _get_appdata_root() -> Path:
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / "VoxCord"
    return Path.home() / "AppData" / "Local" / "VoxCord"


def _load_json_file(path: Path):
    text = path.read_text(encoding="utf-8")
    return json.loads(text)


def kill_existing_voicevox_run_exe() -> None:
    """
    起動前に run.exe を全部終了する。
    0 と 128 は正常扱い。
    """
    if os.name != "nt":
        return

    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "run.exe", "/T"],
            capture_output=True,
            text=True,
            encoding="cp932",
            errors="ignore",
        )

        if result.stdout.strip():
            LOGGER.info("taskkill stdout: %s", result.stdout.strip())
        if result.stderr.strip():
            LOGGER.info("taskkill stderr: %s", result.stderr.strip())

        if result.returncode in (0, 128):
            LOGGER.info("taskkill completed normally (code=%s)", result.returncode)
        else:
            LOGGER.warning("taskkill returned unexpected code=%s", result.returncode)

    except Exception:
        LOGGER.exception("Failed to kill existing run.exe processes")


class TTSEngine:
    """
    VOICEVOX を使う TTS エンジン。

    方針:
    - 起動前に run.exe を全部終了
    - URL は 127.0.0.1 固定
    - Uvicorn running ... を検知したら起動成功とする
    - 話者一覧は speakers.json から読む
    - 再試行しない
    - 自分で起動したプロセスだけ停止する
    """

    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = _resolve_base_dir(base_dir)

        self.voicevox_dir = self.base_dir / "VOICEVOX"
        self.voicevox_exe = self.voicevox_dir / "run.exe"

        self.ffmpeg_dir = self.base_dir / "FFmpeg"
        self.ffmpeg_exe = self.ffmpeg_dir / "ffmpeg.exe"

        self.voicevox_url = "http://127.0.0.1:50021"
        self.voicevox_health_url = f"{self.voicevox_url}/speakers"

        self._voicevox_process: Optional[subprocess.Popen] = None
        self._started = False
        self._owns_process = False
        self._start_attempted = False

        self._start_lock = threading.Lock()
        self._request_lock = threading.Lock()

        self._appdata_root = _get_appdata_root()
        self._temp_dir = self._appdata_root / "temp"
        self._temp_dir.mkdir(parents=True, exist_ok=True)

        self._log_dir = self._appdata_root / "logs"
        self._log_dir.mkdir(parents=True, exist_ok=True)

        self._voicevox_proc_log = self._log_dir / "voicevox_process.log"
        self._voicevox_run_log = self._log_dir / "voicevox_run.log"

        self._stdout_thread: Optional[threading.Thread] = None
        self._stderr_thread: Optional[threading.Thread] = None
        self._log_tail: list[str] = []
        self._log_tail_limit = 200
        self._ready_event = threading.Event()

        LOGGER.info(
            "TTSEngine init: base_dir=%s voicevox=%s ffmpeg=%s temp=%s",
            self.base_dir,
            self.voicevox_exe,
            self.ffmpeg_exe,
            self._temp_dir,
        )

    def _now(self) -> str:
        return time.strftime("%Y-%m-%dT%H:%M:%S%z")

    def _append_process_log(self, line: str) -> None:
        try:
            with self._voicevox_proc_log.open("a", encoding="utf-8", errors="replace") as f:
                f.write(f"{self._now()} {line}\n")
        except Exception:
            pass

    def _append_run_log(self, line: str) -> None:
        try:
            with self._voicevox_run_log.open("a", encoding="utf-8", errors="replace") as f:
                f.write(f"{self._now()} {line}\n")
        except Exception:
            pass

    def _tail_add(self, line: str) -> None:
        self._log_tail.append(line)
        if len(self._log_tail) > self._log_tail_limit:
            self._log_tail = self._log_tail[-self._log_tail_limit :]

    def _dump_tail(self) -> None:
        if not self._log_tail:
            return
        self._append_process_log("----- VOICEVOX log tail begin -----")
        for line in self._log_tail[-40:]:
            self._append_process_log(line.rstrip("\n"))
        self._append_process_log("----- VOICEVOX log tail end -----")

    def _is_process_alive(self) -> bool:
        try:
            return self._voicevox_process is not None and self._voicevox_process.poll() is None
        except Exception:
            return False

    def _health_check(self, timeout: float = 2.0) -> bool:
        """
        /speakers にアクセスできるかだけ確認する。
        JSONの中身は使わず、HTTP 200 を見たら成功とする。
        """
        try:
            resp = requests.get(
                self.voicevox_health_url,
                timeout=timeout,
                headers={"Connection": "close"},
            )
            return resp.status_code == 200
        except Exception:
            return False

    def _request(
        self,
        method: str,
        url: str,
        *,
        params=None,
        json_body=None,
        timeout: float = 30.0,
    ) -> requests.Response:
        with self._request_lock:
            resp = requests.request(
                method,
                url,
                params=params,
                json=json_body,
                timeout=timeout,
                headers={"Connection": "close"},
            )
            resp.raise_for_status()
            return resp

    def _mark_ready_from_line(self, text: str) -> None:
        if (
            "Uvicorn running on http://localhost:50021" in text
            or "Uvicorn running on http://127.0.0.1:50021" in text
        ):
            if not self._ready_event.is_set():
                self._ready_event.set()
                LOGGER.info("VOICEVOX ready marker detected from log line")
                self._append_process_log("VOICEVOX ready marker detected from log line")

    def _stream_reader(self, stream, tag: str) -> None:
        try:
            if stream is None:
                return
            for line in iter(stream.readline, ""):
                if line == "":
                    break
                text = line.rstrip("\n")
                self._tail_add(f"[{tag}] {text}")
                self._append_run_log(f"[{tag}] {text}")
                self._mark_ready_from_line(text)
                LOGGER.info("VOICEVOX %s: %s", tag, text)
        except Exception as e:
            self._tail_add(f"[{tag}] reader error: {e}")
            self._append_run_log(f"[{tag}] reader error: {e}")
            LOGGER.exception("VOICEVOX %s reader failed", tag)

    def _start_log_threads(self) -> None:
        if self._voicevox_process is None:
            return

        if self._voicevox_process.stdout is not None:
            self._stdout_thread = threading.Thread(
                target=self._stream_reader,
                args=(self._voicevox_process.stdout, "stdout"),
                name="VOICEVOX-stdout",
                daemon=True,
            )
            self._stdout_thread.start()

        if self._voicevox_process.stderr is not None:
            self._stderr_thread = threading.Thread(
                target=self._stream_reader,
                args=(self._voicevox_process.stderr, "stderr"),
                name="VOICEVOX-stderr",
                daemon=True,
            )
            self._stderr_thread.start()

    def _start_popen(self) -> None:
        popen_kwargs = {
            "cwd": str(self.voicevox_dir),
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "bufsize": 1,
        }

        # Windowsでは run.exe のコンソールを非表示にする
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            popen_kwargs["startupinfo"] = startupinfo
            popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        self._voicevox_process = subprocess.Popen(
            [str(self.voicevox_exe)],
            **popen_kwargs,
        )

        self._owns_process = True
        self._started = False

        LOGGER.info(
            "VOICEVOX process started: pid=%s",
            self._voicevox_process.pid,
        )
        self._append_process_log(
            f"VOICEVOX process started: pid={self._voicevox_process.pid}"
        )
        self._append_run_log(
            f"=== launch pid={self._voicevox_process.pid} ==="
        )

        self._start_log_threads()

    def _wait_until_ready(self, timeout: float) -> bool:
        deadline = time.monotonic() + timeout

        while time.monotonic() < deadline:
            if self._voicevox_process is not None:
                rc = self._voicevox_process.poll()
                if rc is not None:
                    LOGGER.error("VOICEVOX exited early: returncode=%s", rc)
                    self._append_process_log(f"VOICEVOX exited early: returncode={rc}")
                    self._dump_tail()
                    return False

            if self._ready_event.is_set():
                self._started = True
                LOGGER.info("VOICEVOX is ready (log marker): %s", self.voicevox_health_url)
                self._append_process_log(f"VOICEVOX ready: {self.voicevox_health_url}")
                return True

            if self._health_check(timeout=1.5):
                self._started = True
                LOGGER.info("VOICEVOX is ready (health check): %s", self.voicevox_health_url)
                self._append_process_log(f"VOICEVOX ready: {self.voicevox_health_url}")
                return True

            time.sleep(0.5)

        LOGGER.error("VOICEVOX startup timeout after %.2fs", timeout)
        self._append_process_log(f"VOICEVOX startup timeout after {timeout:.2f}s")
        self._dump_tail()
        return False

    def _ensure_voicevox_running(self, timeout: float = 90.0) -> None:
        """
        合成時に呼ぶ。
        既に起動済みなら何もしない。
        """
        if self._started and self._is_process_alive():
            return

        if self._health_check(timeout=2.0):
            self._started = True
            self._owns_process = False
            self._voicevox_process = None
            LOGGER.info("VOICEVOX already running on %s", self.voicevox_health_url)
            self._append_process_log(f"VOICEVOX already running on {self.voicevox_health_url}")
            return

        if not self._start_attempted:
            self.start_voicevox(timeout=timeout)
            return

        raise RuntimeError("VOICEVOX は既に起動試行済みです。再試行はしません")

    def start_voicevox(self, timeout: float = 90.0) -> None:
        """
        VOICEVOX を起動し、Uvicorn の起動ログを見たら成功とする。

        方針:
        - 起動前に既存の run.exe をすべて終了する
        - 失敗したら同一インスタンスで再試行しない
        - URL は 127.0.0.1 固定
        - ready ログを見たら即成功扱い
        """
        with self._start_lock:
            if self._started and self._is_process_alive() and self._health_check(timeout=2.0):
                return

            if self._health_check(timeout=2.0):
                self._started = True
                self._owns_process = False
                self._voicevox_process = None
                LOGGER.info("VOICEVOX already running on %s", self.voicevox_health_url)
                self._append_process_log(f"VOICEVOX already running on {self.voicevox_health_url}")
                return

            if self._start_attempted:
                raise RuntimeError("VOICEVOX は既に起動試行済みです。再試行はしません")

            self._ready_event.clear()

            kill_existing_voicevox_run_exe()
            time.sleep(1.5)

            if self._health_check(timeout=2.0):
                self._started = True
                self._owns_process = False
                self._voicevox_process = None
                LOGGER.info("VOICEVOX already running on %s", self.voicevox_health_url)
                self._append_process_log(f"VOICEVOX already running on {self.voicevox_health_url}")
                return

            if not self.voicevox_exe.exists():
                raise FileNotFoundError(f"VOICEVOX run.exe が見つかりません: {self.voicevox_exe}")

            if not self.voicevox_dir.exists():
                raise FileNotFoundError(f"VOICEVOX フォルダが見つかりません: {self.voicevox_dir}")

            self._start_attempted = True
            LOGGER.info("Starting VOICEVOX: %s", self.voicevox_exe)
            self._append_process_log(f"Starting VOICEVOX: {self.voicevox_exe}")

            try:
                self._start_popen()
            except Exception:
                LOGGER.exception("VOICEVOX Popen failed")
                self._append_process_log("VOICEVOX Popen failed")
                self._voicevox_process = None
                self._started = False
                self._owns_process = False
                raise

            if self._wait_until_ready(timeout=timeout):
                return

            raise RuntimeError("VOICEVOX の起動確認に失敗しました")

    def stop_voicevox(self) -> None:
        """
        自分で起動した VOICEVOX だけ止める。
        """
        if not self._owns_process:
            self._started = False
            self._voicevox_process = None
            return

        if self._voicevox_process:
            try:
                LOGGER.info("Stopping VOICEVOX: pid=%s", self._voicevox_process.pid)
                self._append_process_log(f"Stopping VOICEVOX: pid={self._voicevox_process.pid}")
                self._voicevox_process.terminate()
                try:
                    self._voicevox_process.wait(timeout=5)
                except Exception:
                    LOGGER.warning("VOICEVOX did not terminate in time")
                    self._append_process_log("VOICEVOX did not terminate in time")
            except Exception:
                LOGGER.exception("VOICEVOX stop failed")
                self._append_process_log("VOICEVOX stop failed")
            finally:
                self._voicevox_process = None
                self._started = False
                self._owns_process = False
        else:
            self._started = False
            self._owns_process = False

    def synthesize(self, text: str, speaker: int, speed: float = 1.0) -> str:
        """
        wavファイルを生成してパスを返す。
        起動済みなら再起動せず、そのまま使う。
        """
        self._ensure_voicevox_running()

        audio_query_url = f"{self.voicevox_url}/audio_query"
        synthesis_url = f"{self.voicevox_url}/synthesis"

        try:
            LOGGER.info("VOICEVOX audio_query: speaker=%s text_len=%s", speaker, len(text))
            self._append_process_log(f"audio_query speaker={speaker} text_len={len(text)}")

            query_resp = self._request(
                "POST",
                audio_query_url,
                params={"text": text, "speaker": speaker},
                timeout=30.0,
            )
            query = query_resp.json()
            query["speedScale"] = float(speed)

            LOGGER.info("VOICEVOX synthesis: speaker=%s speed=%s", speaker, speed)
            self._append_process_log(f"synthesis speaker={speaker} speed={speed}")

            wav_resp = self._request(
                "POST",
                synthesis_url,
                params={"speaker": speaker},
                json_body=query,
                timeout=60.0,
            )

        except Exception as e:
            LOGGER.exception("VOICEVOX synthesis failed")
            self._append_process_log(f"synthesis failed: {e}")
            raise RuntimeError(f"VOICEVOX 合成に失敗しました: {e}") from e

        tmp_wav = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".wav",
            dir=str(self._temp_dir),
        )
        try:
            tmp_wav.write(wav_resp.content)
        finally:
            tmp_wav.close()

        LOGGER.info("Generated wav: %s", tmp_wav.name)
        self._append_process_log(f"Generated wav: {tmp_wav.name}")
        return tmp_wav.name

    def convert_for_discord(self, wav_path: str) -> str:
        """
        48kHz stereo PCM に変換
        """
        if not self.ffmpeg_exe.exists():
            raise FileNotFoundError(f"FFmpeg が見つかりません: {self.ffmpeg_exe}")

        src = Path(wav_path)
        output_path = src.with_name(f"{src.stem}_48k.wav")

        cmd = [
            str(self.ffmpeg_exe),
            "-y",
            "-i",
            str(src),
            "-ar",
            "48000",
            "-ac",
            "2",
            str(output_path),
        ]

        kwargs = {
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
        }

        LOGGER.info("Converting wav for discord: %s -> %s", src, output_path)
        self._append_process_log(f"Converting wav for discord: {src} -> {output_path}")

        result = subprocess.run(cmd, **kwargs)
        if result.returncode != 0:
            self._append_process_log(f"FFmpeg failed rc={result.returncode}")
            raise RuntimeError("FFmpeg 変換に失敗しました")

        return str(output_path)

    def play_local(self, wav_path: str) -> None:
        os.startfile(wav_path)

    def cleanup(self, *paths) -> None:
        for p in paths:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
                    LOGGER.debug("Deleted temp wav: %s", p)
            except Exception:
                LOGGER.exception("Failed to delete temp wav: %s", p)

    def load_speakers(self) -> List[Dict[str, int]]:
        """
        話者一覧は local の speakers.json だけを読む。
        API の JSON は使わない。
        """
        path = self.base_dir / "speakers.json"
        if not path.exists():
            raise FileNotFoundError(f"speakers.json が見つかりません: {path}")

        data = _load_json_file(path)
        if not isinstance(data, list):
            raise ValueError("speakers.json の形式が不正です (list ではありません)")

        speakers: List[Dict[str, int]] = []
        for s in data:
            if not isinstance(s, dict):
                continue
            name = str(s.get("name", "Unknown"))
            styles = s.get("styles", [])
            if not isinstance(styles, list):
                continue
            for style in styles:
                if not isinstance(style, dict):
                    continue
                try:
                    speakers.append(
                        {
                            "name": f"{name} ({style.get('name', 'Unknown')})",
                            "id": int(style["id"]),
                        }
                    )
                except Exception:
                    continue

        return speakers

    def get_speakers(self) -> List[Dict[str, int]]:
        """
        話者取得は speakers.json だけを使う。
        """
        return self.load_speakers()

    def close(self) -> None:
        try:
            self.stop_voicevox()
        finally:
            pass

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass
