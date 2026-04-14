from __future__ import annotations

import os
import subprocess
import tempfile
import time
import logging
from pathlib import Path
from typing import List, Dict, Optional

import requests

LOGGER = logging.getLogger(__name__)

import json
from pathlib import Path
import sys

def load_speakers():
    """
    exeと同じ場所の speakers.json を読み込む
    """
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    path = base_dir / "speakers.json"

    if not path.exists():
        raise FileNotFoundError(f"speakers.json が見つかりません: {path}")

    text = path.read_text(encoding="utf-8")
    data = json.loads(text)

    return data

def _get_appdata_root() -> Path:
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / "VoxCord"
    return Path.home() / "AppData" / "Local" / "VoxCord"


def _get_base_dir(base_dir: str) -> Path:
    return Path(base_dir).resolve()


class TTSEngine:
    def __init__(self, base_dir: str):
        self.base_dir = _get_base_dir(base_dir)

        # exe と同じ場所に外部フォルダを置く想定
        self.voicevox_dir = self.base_dir / "VOICEVOX"
        self.voicevox_exe = self.voicevox_dir / "run.exe"

        self.ffmpeg_dir = self.base_dir / "FFmpeg"
        self.ffmpeg_exe = self.ffmpeg_dir / "ffmpeg.exe"

        self.voicevox_url = "http://127.0.0.1:50021"

        self._voicevox_process: Optional[subprocess.Popen] = None
        self._started = False
        self._session = requests.Session()

        # 一時 wav は AppData 配下へ
        self._appdata_root = _get_appdata_root()
        self._temp_dir = self._appdata_root / "temp"
        self._temp_dir.mkdir(parents=True, exist_ok=True)

        LOGGER.info(
            "TTSEngine init: base_dir=%s voicevox=%s ffmpeg=%s temp=%s",
            self.base_dir,
            self.voicevox_exe,
            self.ffmpeg_exe,
            self._temp_dir,
        )

    def _subprocess_kwargs(self) -> dict:
        """
        Windowsでコンソールウィンドウを出さないための設定。
        """
        kwargs = {
            "cwd": str(self.voicevox_dir),
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
        }

        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            kwargs["startupinfo"] = startupinfo

        return kwargs

    # =============================
    # VOICEVOX 起動
    # =============================
    def start_voicevox(self) -> None:
        if self._started and self._voicevox_process and self._voicevox_process.poll() is None:
            LOGGER.debug("VOICEVOX already started")
            return

        if not self.voicevox_exe.exists():
            raise FileNotFoundError(f"VOICEVOX run.exe が見つかりません: {self.voicevox_exe}")

        LOGGER.info("Starting VOICEVOX: %s", self.voicevox_exe)

        self._voicevox_process = subprocess.Popen(
            [str(self.voicevox_exe)],
            **self._subprocess_kwargs(),
        )

        # 起動待ち
        ready = False
        for _ in range(30):
            try:
                r = self._session.get(f"{self.voicevox_url}/version", timeout=1.0)
                if r.ok:
                    ready = True
                    break
            except Exception:
                time.sleep(1)

        if not ready:
            raise RuntimeError("VOICEVOX の起動確認に失敗しました")

        self._started = True
        LOGGER.info("VOICEVOX started")

    # =============================
    # VOICEVOX 停止
    # =============================
    def stop_voicevox(self) -> None:
        if self._voicevox_process:
            try:
                LOGGER.info("Stopping VOICEVOX")
                self._voicevox_process.terminate()
                try:
                    self._voicevox_process.wait(timeout=5)
                except Exception:
                    self._voicevox_process.kill()
            except Exception:
                LOGGER.exception("VOICEVOX stop failed")
            finally:
                self._voicevox_process = None
                self._started = False

    # =============================
    # 音声生成
    # =============================
    def synthesize(self, text: str, speaker: int, speed: float = 1.0) -> str:
        """
        wavファイルを生成してパスを返す
        """
        self.start_voicevox()

        audio_query_url = f"{self.voicevox_url}/audio_query"
        synthesis_url = f"{self.voicevox_url}/synthesis"

        try:
            query_resp = self._session.post(
                audio_query_url,
                params={"text": text, "speaker": speaker},
                timeout=30.0,
            )
            query_resp.raise_for_status()
            query = query_resp.json()

            query["speedScale"] = float(speed)

            wav_resp = self._session.post(
                synthesis_url,
                params={"speaker": speaker},
                json=query,
                timeout=60.0,
            )
            wav_resp.raise_for_status()

        except Exception as e:
            LOGGER.exception("VOICEVOX synthesis failed")
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
        return tmp_wav.name

    # =============================
    # FFmpeg 変換（discord用）
    # =============================
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
            "-i", str(src),
            "-ar", "48000",
            "-ac", "2",
            str(output_path),
        ]

        kwargs = {
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
        }
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        LOGGER.info("Converting wav for discord: %s -> %s", src, output_path)
        result = subprocess.run(cmd, **kwargs)
        if result.returncode != 0:
            raise RuntimeError("FFmpeg 変換に失敗しました")

        return str(output_path)

    # =============================
    # 音声再生（ローカル確認用）
    # =============================
    def play_local(self, wav_path: str) -> None:
        """
        Windows の標準再生
        """
        os.startfile(wav_path)

    # =============================
    # wav削除
    # =============================
    def cleanup(self, *paths) -> None:
        for p in paths:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
                    LOGGER.debug("Deleted temp wav: %s", p)
            except Exception:
                LOGGER.exception("Failed to delete temp wav: %s", p)

    # =============================
    # 話者一覧取得
    # =============================
    def get_speakers(self) -> List[Dict[str, int]]:
        """
        VOICEVOXの話者一覧を取得
        """
        self.start_voicevox()

        try:
            res = self._session.get(f"{self.voicevox_url}/speakers", timeout=30.0)
            res.raise_for_status()
            data = res.json()
        except Exception as e:
            LOGGER.exception("Failed to get speakers")
            raise RuntimeError(f"話者一覧の取得に失敗しました: {e}") from e

        speakers: List[Dict[str, int]] = []
        for s in data:
            for style in s.get("styles", []):
                speakers.append(
                    {
                        "name": f"{s.get('name', 'Unknown')} ({style.get('name', 'Unknown')})",
                        "id": int(style["id"]),
                    }
                )

        return speakers

    def close(self) -> None:
        """
        明示的な後始末
        """
        try:
            self.stop_voicevox()
        finally:
            try:
                self._session.close()
            except Exception:
                pass

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass