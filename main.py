# main.py — GUI にログと状態を確実に出す強化版

from __future__ import annotations

import sys
import os
import logging
import logging.handlers
import threading
import asyncio
import traceback
from pathlib import Path
from types import SimpleNamespace
from typing import Optional
import nacl
import nacl.secret
import nacl.utils
import _cffi_backend  # PyNaCl の依存先を先に読み込む（PyInstaller 対策）
import cffi  # PyInstaller 対策

_ = nacl  # PyNaCl を先に読み込んで、音声機能の有無を早めに確定させる
_ = nacl.secret  # secret も使うので先に読み込む
_ = nacl.utils  # utils も使うので先に読み込む
_ = _cffi_backend  # PyInstaller 対策
_ = cffi  # PyInstaller 対策

import discord.opus
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
opus_path = os.path.join(base_dir, "opus.dll")

print("=== START ===")

if os.path.exists(opus_path):
    discord.opus.load_opus(opus_path)

# -------------------------
# 実行元 / AppData
# -------------------------
if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
    MEIPASS_DIR = getattr(sys, "_MEIPASS", None)
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
    MEIPASS_DIR = None

APPDATA_ROOT = Path(os.getenv("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")) / "VoxCord"
APPDATA_ROOT.mkdir(parents=True, exist_ok=True)

# 相対パスで書くログや一時ファイルを AppData に寄せる
os.chdir(APPDATA_ROOT)

# PyInstaller / DLL 対策
try:
    if MEIPASS_DIR and os.path.exists(MEIPASS_DIR):
        os.add_dll_directory(MEIPASS_DIR)
except Exception:
    pass

for _p in (
    str(SCRIPT_DIR),
    str(SCRIPT_DIR / "_internal"),
    str(SCRIPT_DIR / "nacl"),
):
    if os.path.exists(_p):
        try:
            os.add_dll_directory(_p)
        except Exception:
            pass

# PyNaCl を先に読み込んで、音声機能の有無を早めに確定させる
try:
    import nacl  # noqa: F401
    import nacl.bindings  # noqa: F401
except Exception:
    pass

# PySide6
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QGuiApplication, QIcon

# local modules
sys.path.insert(0, str(SCRIPT_DIR))

from config_manager import ConfigManager
import discord_service as ds_mod
from discord_service import DiscordService
from gui import VoxCordGUI

# tts_engine (optional)
try:
    import tts_engine as raw_tts_module
except Exception:
    raw_tts_module = None


# -------------------------
# Logging 設定（ファイル+コンソール）
# -------------------------
LOG_DIR = APPDATA_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)


def setup_logging(level: int = logging.INFO):
    """RotatingFile + Console を設定"""
    log_file = LOG_DIR / "latest.log"
    root = logging.getLogger()
    root.setLevel(level)

    if root.handlers:
        for h in list(root.handlers):
            root.removeHandler(h)

    fmt = "%(asctime)s %(levelname)-5s [%(threadName)s] %(name)s:%(lineno)d - %(message)s"
    datefmt = "%Y-%m-%dT%H:%M:%S%z"
    formatter = logging.Formatter(fmt=fmt, datefmt=datefmt)

    fh = logging.handlers.RotatingFileHandler(
        filename=str(log_file), maxBytes=1_000_000, backupCount=5, encoding="utf-8"
    )
    fh.setFormatter(formatter)
    fh.setLevel(level)
    root.addHandler(fh)

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(formatter)
    ch.setLevel(level)
    root.addHandler(ch)

    logging.getLogger("asyncio").setLevel(logging.WARNING)
    logging.getLogger("discord").setLevel(logging.INFO)

    logging.info("Logging initialized. file=%s", log_file)


# 先にログを有効化
setup_logging(logging.INFO)


# -------------------------
# Qt に安全にログを流す handler
# -------------------------
class QtLogHandler(logging.Handler):
    """
    ログレコードを GUI のログ表示に送るハンドラ。
    emit はどのスレッドでも呼ばれるため、QTimer.singleShot を使い
    Qt イベントスレッドで GUI.append_log を呼ぶ（スレッドセーフ）。
    """

    def __init__(self, gui: VoxCordGUI):
        super().__init__()
        self.gui = gui
        self.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    def emit(self, record: logging.LogRecord):
        try:
            msg = self.format(record)
            QTimer.singleShot(0, lambda m=msg: self._safe_append(m, record))
        except Exception:
            try:
                print("QtLogHandler.emit error", file=sys.stderr)
                print(traceback.format_exc(), file=sys.stderr)
            except Exception:
                pass

    def _safe_append(self, msg: str, record: logging.LogRecord):
        try:
            self.gui.append_log(msg)
            txt = msg.lower()
            if "ログイン成功" in msg or "discord: ログイン成功" in msg or "discord: login" in txt or "discordservice started" in txt:
                self.gui.set_status("RUNNING")
            elif "tts engine started" in txt or ("voicevox" in txt and "started" in txt):
                self.gui.append_log("[INFO] TTS 起動完了")
            elif record.levelno >= logging.ERROR:
                self.gui.set_status("ERROR")
        except Exception:
            try:
                print("Failed to append log to GUI:", msg)
            except Exception:
                pass


# -------------------------
# AppController（バックグラウンドループ + start/stop）
# -------------------------
class AppController:
    def __init__(self, config_path: str = "config.json", base_dir: str = "."):
        self.base_dir = Path(base_dir).resolve()
        logging.info("AppController init: base_dir=%s config=%s", self.base_dir, config_path)
        try:
            self.config = ConfigManager(config_path)
            logging.info("Loaded config: %s", str(self.config_path_summary()))
        except Exception:
            logging.exception("Config load failed, creating new config")
            self.config = ConfigManager(config_path)

        self.loop: asyncio.AbstractEventLoop = asyncio.new_event_loop()
        self._loop_thread = threading.Thread(target=self._run_loop, name="AsyncLoopThread", daemon=True)
        self._loop_thread.start()

        self._discord_service: Optional[DiscordService] = None
        self._tts_instance = None
        self._tts_started = False

        self._start_future = None
        self._stop_future = None

    def config_path_summary(self):
        try:
            d = self.config.to_dict()
            keys = list(d.keys())
            return f"keys={keys}"
        except Exception:
            return "unavailable"

    def _run_loop(self):
        asyncio.set_event_loop(self.loop)

        def loop_ex(loop, context):
            msg = context.get("message", "<no message>")
            exc = context.get("exception")
            logging.error("Unhandled asyncio exception: %s; exc=%s", msg, exc)

        self.loop.set_exception_handler(loop_ex)

        logging.info("Asyncio loop thread start")
        try:
            self.loop.run_forever()
        finally:
            logging.info("Asyncio loop thread exiting")
            try:
                self.loop.close()
            except Exception:
                pass

    def _run_coro_threadsafe(self, coro, desc: str = "<coro>"):
        logging.debug("Scheduling coro: %s", desc)
        fut = asyncio.run_coroutine_threadsafe(coro, self.loop)

        def _done(f):
            try:
                exc = f.exception()
                if exc:
                    logging.error("Task %s failed: %s", desc, exc)
                else:
                    logging.info("Task %s completed", desc)
            except Exception:
                logging.exception("Exception in done callback for %s", desc)

        fut.add_done_callback(_done)
        return fut

    async def _start_tts_engine_coroutine(self) -> bool:
        if self._tts_started:
            logging.info("TTS already started")
            return True
        if raw_tts_module is None or not hasattr(raw_tts_module, "TTSEngine"):
            logging.warning("TTS engine unavailable")
            return False
        try:
            self._tts_instance = raw_tts_module.TTSEngine(str(self.base_dir))
            logging.info("Instantiated TTSEngine")
            await asyncio.to_thread(self._tts_instance.start_voicevox)
            self._tts_started = True
            logging.info("TTS engine started")

            def make_adapter(engine):
                async def synth(text, speaker, speed=1.0):
                    path = await asyncio.to_thread(engine.synthesize, text, int(speaker), float(speed))
                    if hasattr(engine, "convert_for_discord"):
                        try:
                            converted = await asyncio.to_thread(engine.convert_for_discord, path)
                            return converted
                        except Exception:
                            logging.exception("convert_for_discord failed")
                            return path
                    return path
                return synth

            ds_mod.tts_engine = SimpleNamespace(synthesize_wav=make_adapter(self._tts_instance))
            logging.info("TTS adapter installed into discord_service")
            return True
        except Exception:
            logging.exception("Failed to start TTS engine")
            return False

    async def _stop_tts_engine_coroutine(self):
        if not self._tts_started or not self._tts_instance:
            logging.debug("TTS not running")
            return
        try:
            await asyncio.to_thread(self._tts_instance.stop_voicevox)
            logging.info("TTS engine stopped")
        except Exception:
            logging.exception("Error stopping TTS engine")
        finally:
            self._tts_started = False
            self._tts_instance = None
            ds_mod.tts_engine = None

    async def _start_discord_coroutine(self):
        if self._discord_service is not None:
            logging.info("DiscordService already exists")
            return
        try:
            self._discord_service = DiscordService(self.config)
            await self._discord_service.start()
            logging.info("DiscordService started")
        except Exception:
            logging.exception("DiscordService start failed")
            if self._discord_service:
                try:
                    await self._discord_service.stop()
                except Exception:
                    logging.exception("DiscordService cleanup failed")
            self._discord_service = None
            raise

    async def _stop_discord_coroutine(self):
        if self._discord_service is None:
            logging.debug("DiscordService not running")
            return
        try:
            await self._discord_service.stop()
            logging.info("DiscordService stopped")
        except Exception:
            logging.exception("DiscordService stop failed")
        finally:
            self._discord_service = None

    async def start_all(self):
        logging.info("start_all: sequence begin")
        ok = await self._start_tts_engine_coroutine()
        if not ok:
            logging.warning("TTS failed — abort start")
            return
        await self._start_discord_coroutine()
        logging.info("start_all: sequence end")

    async def stop_all(self):
        logging.info("stop_all: sequence begin")
        await self._stop_discord_coroutine()
        await self._stop_tts_engine_coroutine()
        logging.info("stop_all: sequence end")

    def start_all_async(self):
        fut = self._run_coro_threadsafe(self.start_all(), desc="start_all")
        self._start_future = fut
        return fut

    def stop_all_async(self):
        fut = self._run_coro_threadsafe(self.stop_all(), desc="stop_all")
        self._stop_future = fut
        return fut

    def cleanup_sync(self):
        logging.info("cleanup_sync: scheduling shutdown")
        try:
            fut = self._run_coro_threadsafe(self.stop_all(), desc="cleanup stop_all")
            try:
                fut.result(timeout=8.0)
                logging.info("cleanup_sync: stop_all finished quickly")
            except Exception:
                logging.info("cleanup_sync: stop_all didn't finish quickly; continuing")
        except Exception:
            logging.exception("cleanup_sync: scheduling failed")
        finally:
            def _stop_loop():
                try:
                    self.loop.stop()
                except Exception:
                    pass

            self.loop.call_soon_threadsafe(_stop_loop)
            try:
                if self._loop_thread.is_alive():
                    self._loop_thread.join(timeout=1.0)
            except Exception:
                pass


# -------------------------
# 起動エントリポイント
# -------------------------
def main():
    logging.info("VoxCord main starting")

    try:
        QGuiApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QGuiApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    except Exception:
        logging.debug("High-DPI attributes not applied")

    app = QApplication(sys.argv)

    # アイコン
    icon = None
    for icon_path in (
        SCRIPT_DIR / "logo.ico",
        SCRIPT_DIR / "assets" / "logo.ico",
        SCRIPT_DIR / "assets" / "logo.png",
        SCRIPT_DIR / "logo.png",
    ):
        if icon_path.exists():
            icon = QIcon(str(icon_path))
            break

    if icon is not None and not icon.isNull():
        app.setWindowIcon(icon)

    config_path = str(APPDATA_ROOT / "config.json")
    ctrl = AppController(config_path=config_path, base_dir=str(SCRIPT_DIR))

    gui = VoxCordGUI(ctrl.config)
    if icon is not None and not icon.isNull():
        gui.setWindowIcon(icon)

    qt_handler = QtLogHandler(gui)
    qt_handler.setLevel(logging.INFO)
    logging.getLogger().addHandler(qt_handler)
    gui.append_log("[INFO] アプリを起動しました。Start ボタンを押してください。")

    def attach_future_callbacks(fut, label_on_success="RUNNING"):
        if fut is None:
            return

        def _cb(f):
            try:
                exc = f.exception()
                if exc:
                    QTimer.singleShot(0, lambda: (
                        gui.append_log(f"[ERROR] 起動処理で例外: {exc}"),
                        gui.set_status("ERROR"),
                        QMessageBox.critical(gui, "起動エラー", f"起動に失敗しました:\n{exc}")
                    ))
                else:
                    QTimer.singleShot(0, lambda: (
                        gui.append_log("[INFO] 起動処理が完了しました"),
                        gui.set_status(label_on_success)
                    ))
            except Exception:
                QTimer.singleShot(0, lambda: gui.append_log("[ERROR] Future callback failed"))

        try:
            fut.add_done_callback(lambda f: _cb(f))
        except Exception:
            logging.exception("attach_future_callbacks failed")

    def on_start_clicked():
        logging.info("[INFO] 起動要求を送信しました")
        gui.append_log("[INFO] 接続しました。")
        gui.set_status("CONNECTING")
        fut = ctrl.start_all_async()
        attach_future_callbacks(fut, label_on_success="RUNNING")

    def on_stop_clicked():
        logging.info("[INFO]停止要求を送信しました。")
        gui.append_log("[INFO] 終了しました。")
        gui.set_status("STOPPING")
        fut = ctrl.stop_all_async()
        attach_future_callbacks(fut, label_on_success="STOPPED")

    gui.start_button.clicked.connect(on_start_clicked)
    gui.stop_button.clicked.connect(on_stop_clicked)

    def poll_state():
        try:
            if getattr(ctrl, "_start_future", None) and not getattr(ctrl._start_future, "done", lambda: True)():
                gui.set_status("CONNECTING")
                return
            ds = getattr(ctrl, "_discord_service", None)
            if ds and getattr(ds, "client", None) and getattr(ds.client, "user", None):
                gui.set_status("RUNNING")
                return
            if getattr(ctrl, "_stop_future", None) and not getattr(ctrl._stop_future, "done", lambda: True)():
                gui.set_status("STOPPING")
                return
            gui.set_status("STOPPED")
        except Exception:
            logging.exception("poll_state exception")

    poll_timer = QTimer()
    poll_timer.timeout.connect(poll_state)
    poll_timer.start(2000)

    def _on_about_to_quit():
        logging.info("Qt aboutToQuit: cleanup start")
        gui.append_log("[INFO] アプリ終了: クリーンアップ中")
        ctrl.cleanup_sync()

    app.aboutToQuit.connect(_on_about_to_quit)

    gui.show()
    exit_code = app.exec()
    logging.info("Qt loop exited (code=%s)", exit_code)

    logging.info("Exiting process")
    sys.exit(0)


if __name__ == "__main__":
    main()