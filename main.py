from __future__ import annotations

import asyncio
import logging
import logging.handlers
import os
import sys
import threading
import traceback
from pathlib import Path
from types import SimpleNamespace
from typing import Optional

import discord.opus
import nacl
import nacl.secret
import nacl.utils
import _cffi_backend
import cffi

_ = nacl
_ = nacl.secret
_ = nacl.utils
_ = _cffi_backend
_ = cffi

import discord
import os
import sys

def load_opus():
    base_dir = os.path.dirname(sys.executable)  # exe用
    opus_path = os.path.join(base_dir, "opus.dll")

    if not discord.opus.is_loaded():
        if os.path.exists(opus_path):
            discord.opus.load_opus(opus_path)
            print("Opus loaded:", opus_path)
        else:
            print("Opus not found:", opus_path)

load_opus()

base_dir = os.path.dirname(os.path.abspath(__file__))
opus_path = os.path.join(base_dir, "opus.dll")

from pathlib import Path
import sys

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

print("=== START ===")

import os

os.environ.pop("HTTP_PROXY", None)
os.environ.pop("HTTPS_PROXY", None)
os.environ.pop("http_proxy", None)
os.environ.pop("https_proxy", None)
os.environ.pop("ALL_PROXY", None)
os.environ.pop("all_proxy", None)

import discord
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

opus_path = os.path.join(base_dir, "opus.dll")

if not discord.opus.is_loaded():
    discord.opus.load_opus(opus_path)

print("Opus loaded:", discord.opus.is_loaded())

if os.path.exists(opus_path):
    discord.opus.load_opus(opus_path)

if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
    MEIPASS_DIR = getattr(sys, "_MEIPASS", None)
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
    MEIPASS_DIR = None

APPDATA_ROOT = Path(
    os.getenv("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")
) / "VoxCord"
APPDATA_ROOT.mkdir(parents=True, exist_ok=True)
os.chdir(APPDATA_ROOT)

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

try:
    import nacl  # noqa: F401
    import nacl.bindings  # noqa: F401
except Exception:
    pass

from PySide6.QtCore import QObject, QThread, QTimer, Signal, Slot
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtGui import QIcon

sys.path.insert(0, str(SCRIPT_DIR))

from config_manager import ConfigManager
import discord_service as ds_mod
from discord_service import DiscordService
from gui import VoxCordGUI
from load import LoadingWindow

try:
    import tts_engine as raw_tts_module
except Exception:
    raw_tts_module = None


LOG_DIR = APPDATA_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)


def setup_logging(level: int = logging.INFO):
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
        filename=str(log_file),
        maxBytes=1_000_000,
        backupCount=5,
        encoding="utf-8",
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


setup_logging(logging.INFO)


class _LogRelay(QObject):
    message = Signal(str)


class WidgetLogHandler(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        self._relay = _LogRelay()
        self._relay.message.connect(self._safe_append)

    def emit(self, record: logging.LogRecord):
        try:
            msg = self.format(record)
            self._relay.message.emit(msg)
        except Exception:
            try:
                print("WidgetLogHandler.emit error", file=sys.stderr)
                print(traceback.format_exc(), file=sys.stderr)
            except Exception:
                pass

    @Slot(str)
    def _safe_append(self, msg: str):
        try:
            if self.widget is not None and hasattr(self.widget, "append_log"):
                self.widget.append_log(msg)
        except Exception:
            pass


class AppController:
    def __init__(self, base_dir: str = "."):
        self.base_dir = Path(base_dir).resolve()
        logging.info("AppController init: base_dir=%s", self.base_dir)

        self.config: Optional[ConfigManager] = None

        self.loop: asyncio.AbstractEventLoop = asyncio.new_event_loop()
        self._loop_thread = threading.Thread(
            target=self._run_loop,
            name="AsyncLoopThread",
            daemon=True,
        )
        self._loop_thread.start()

        self._discord_service: Optional[DiscordService] = None
        self._tts_instance = None
        self._tts_started = False

        self._start_future = None
        self._stop_future = None

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

    def load_config(self, config_path: str):
        self.config = ConfigManager(config_path)
        logging.info("Loaded config: %s", self.config_path_summary())
        return self.config

    def config_path_summary(self):
        try:
            if self.config is None:
                return "unavailable"
            d = self.config.to_dict()
            keys = list(d.keys())
            return f"keys={keys}"
        except Exception:
            return "unavailable"

    async def _start_tts_engine_coroutine(self, timeout: float = 90.0) -> bool:
        if self._tts_started:
            logging.info("TTS already started")
            return True

        if raw_tts_module is None or not hasattr(raw_tts_module, "TTSEngine"):
            logging.warning("TTS engine unavailable")
            return False

        try:
            self._tts_instance = raw_tts_module.TTSEngine(str(self.base_dir))
            logging.info("Instantiated TTSEngine")

            await asyncio.to_thread(self._tts_instance.start_voicevox, timeout)
            self._tts_started = True
            logging.info("TTS engine started")

            def make_adapter(engine):
                async def synth(text, speaker, speed=1.0):
                    path = await asyncio.to_thread(
                        engine.synthesize,
                        text,
                        int(speaker),
                        float(speed),
                    )
                    if hasattr(engine, "convert_for_discord"):
                        try:
                            converted = await asyncio.to_thread(engine.convert_for_discord, path)
                            return converted
                        except Exception:
                            logging.exception("convert_for_discord failed")
                            return path
                    return path

                return synth

            ds_mod.tts_engine = SimpleNamespace(
                synthesize_wav=make_adapter(self._tts_instance)
            )
            logging.info("TTS adapter installed into discord_service")
            return True

        except Exception:
            logging.exception("Failed to start TTS engine")
            return False

    def start_tts_blocking(self, timeout: float = 90.0) -> bool:
        if self._tts_started:
            logging.info("TTS already started")
            return True

        fut = self._run_coro_threadsafe(
            self._start_tts_engine_coroutine(timeout=timeout),
            desc="start_tts",
        )
        return fut.result(timeout=120.0)

    def stop_tts_blocking(self):
        async def _runner():
            if not self._tts_started or not self._tts_instance:
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

        fut = self._run_coro_threadsafe(_runner(), desc="stop_tts")
        fut.result(timeout=30.0)

    async def _start_discord_coroutine(self):
        if self._discord_service is not None:
            logging.info("DiscordService already exists")
            return
        if self.config is None:
            raise RuntimeError("config is not loaded")
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
        ok = await self._start_tts_engine_coroutine(timeout=90.0)
        if not ok:
            logging.warning("TTS failed — abort start")
            return
        await self._start_discord_coroutine()
        logging.info("start_all: sequence end")

    async def stop_all(self):
        logging.info("stop_all: sequence begin")
        await self._stop_discord_coroutine()
        try:
            self.stop_tts_blocking()
        except Exception:
            logging.exception("stop_tts failed")
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
            try:
                self.loop.call_soon_threadsafe(self.loop.stop)
            except Exception:
                pass
            try:
                if self._loop_thread.is_alive():
                    self._loop_thread.join(timeout=1.0)
            except Exception:
                pass


class StartupWorker(QObject):
    status = Signal(str)
    detail = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, controller: AppController, config_path: str):
        super().__init__()
        self.controller = controller
        self.config_path = config_path

    @Slot()
    def run(self):
        try:
            self.status.emit("VOICEVOX を起動しています…")
            self.detail.emit("既存の run.exe を終了してから起動しています")

            ok = self.controller.start_tts_blocking(timeout=90.0)
            if not ok:
                raise RuntimeError("VOICEVOX の起動に失敗しました")

            self.status.emit("設定を読み込んでいます…")
            self.detail.emit("config.json を読み込んでいます")
            self.controller.load_config(self.config_path)

            self.status.emit("起動準備が完了しました")
            self.detail.emit("メイン画面を開きます")
            self.finished.emit(self.controller)

        except Exception as e:
            logging.exception("StartupWorker failed")
            self.failed.emit(str(e))


class StartupUiBridge(QObject):
    def __init__(self, app: QApplication, loading: LoadingWindow, thread: QThread, state, icon, script_dir: Path):
        super().__init__()
        self.app = app
        self.loading = loading
        self.thread = thread
        self.state = state
        self.icon = icon
        self.script_dir = script_dir

    @Slot(str)
    def on_status(self, text: str):
        try:
            self.loading.set_status(text)
            self.loading.append_log(f"[INFO] {text}")
        except Exception:
            pass

    @Slot(str)
    def on_detail(self, text: str):
        try:
            self.loading.set_detail(text)
            self.loading.append_log(f"  {text}")
        except Exception:
            pass

    @Slot(str)
    def on_failed(self, message: str):
        try:
            self.loading.set_status("起動に失敗しました")
            self.loading.set_detail(message)
            self.loading.append_log(f"[ERROR] {message}")
            QMessageBox.critical(self.loading, "起動エラー", message)
        finally:
            try:
                self.thread.quit()
                self.thread.wait(1000)
            except Exception:
                pass
            try:
                self.app.quit()
            except Exception:
                pass

    @Slot(object)
    def on_finished(self, ctrl: AppController):
        try:
            self.loading.set_status("完了")
            self.loading.set_detail("メイン画面を表示しています")
            self.loading.append_log("[INFO] 起動完了")

            if self.state.loading_log_handler is not None:
                try:
                    logging.getLogger().removeHandler(self.state.loading_log_handler)
                except Exception:
                    pass

            gui = VoxCordGUI(ctrl.config)
            if self.icon is not None and not self.icon.isNull():
                gui.setWindowIcon(self.icon)

            gui_log_handler = WidgetLogHandler(gui)
            gui_log_handler.setLevel(logging.INFO)
            logging.getLogger().addHandler(gui_log_handler)

            self.state.gui = gui
            self.state.gui_log_handler = gui_log_handler

            gui.append_log("[INFO] アプリを起動しました。Start ボタンを押してください。")

            def attach_future_callbacks(fut, label_on_success="RUNNING"):
                if fut is None:
                    return

                def _cb(f):
                    try:
                        exc = f.exception()
                        if exc:
                            QMessageBox.critical(gui, "起動エラー", f"起動に失敗しました:\n{exc}")
                            gui.append_log(f"[ERROR] 起動処理で例外: {exc}")
                            gui.set_status("ERROR")
                        else:
                            gui.append_log("[INFO] 起動処理が完了しました")
                            gui.set_status(label_on_success)
                    except Exception:
                        logging.exception("attach_future_callbacks failed")

                try:
                    fut.add_done_callback(lambda f: QTimer.singleShot(0, lambda: _cb(f)))
                except Exception:
                    logging.exception("attach_future_callbacks failed")

            def on_start_clicked():
                logging.info("[INFO] 起動要求を送信しました")
                gui.append_log("[INFO] 接続しました。")
                gui.set_status("CONNECTING")
                fut = ctrl.start_all_async()
                attach_future_callbacks(fut, label_on_success="RUNNING")

            def on_stop_clicked():
                logging.info("[INFO] 停止要求を送信しました")
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

            poll_timer = QTimer(gui)
            poll_timer.timeout.connect(poll_state)
            poll_timer.start(2000)

            def _on_about_to_quit():
                logging.info("Qt aboutToQuit: cleanup start")
                gui.append_log("[INFO] アプリ終了: クリーンアップ中")
                ctrl.cleanup_sync()

            self.app.aboutToQuit.connect(_on_about_to_quit)

            self.loading.close()
            gui.show()
            gui.raise_()
            gui.activateWindow()

        except Exception as e:
            logging.exception("Failed to open main GUI")
            QMessageBox.critical(self.loading, "起動エラー", str(e))
            self.app.quit()
        finally:
            try:
                self.thread.quit()
                self.thread.wait(1000)
            except Exception:
                pass


def main():
    logging.info("VoxCord main starting")

    app = QApplication(sys.argv)

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

    loading = LoadingWindow()
    if icon is not None and not icon.isNull():
        loading.setWindowIcon(icon)

    loading.show()
    loading.raise_()
    loading.activateWindow()

    controller = AppController(base_dir=str(SCRIPT_DIR))

    state = SimpleNamespace(
        controller=controller,
        gui=None,
        loading=loading,
        loading_log_handler=None,
        gui_log_handler=None,
        worker=None,
        thread=None,
    )

    state.loading_log_handler = WidgetLogHandler(loading)
    state.loading_log_handler.setLevel(logging.INFO)
    logging.getLogger().addHandler(state.loading_log_handler)

    config_path = str(APPDATA_ROOT / "config.json")

    worker = StartupWorker(controller=controller, config_path=config_path)
    thread = QThread()
    worker.moveToThread(thread)

    state.worker = worker
    state.thread = thread

    bridge = StartupUiBridge(
        app=app,
        loading=loading,
        thread=thread,
        state=state,
        icon=icon,
        script_dir=SCRIPT_DIR,
    )

    worker.status.connect(bridge.on_status)
    worker.detail.connect(bridge.on_detail)
    worker.failed.connect(bridge.on_failed)
    worker.finished.connect(bridge.on_finished)

    thread.started.connect(worker.run)

    def start_worker():
        loading.set_status("起動中…")
        loading.set_detail("VOICEVOX と設定を準備しています")
        thread.start()

    QTimer.singleShot(0, start_worker)

    exit_code = app.exec()
    logging.info("Qt loop exited (code=%s)", exit_code)

    try:
        if state.controller is not None:
            state.controller.cleanup_sync()
    except Exception:
        pass

    logging.info("Exiting process")
    sys.exit(0)


if __name__ == "__main__":
    main()
