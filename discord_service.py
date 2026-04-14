"""
DiscordService - Discord 接続 / メッセージ監視 / 再生キュー管理

主な機能:
 - Bot接続（非同期）
 - 指定された text_channel_id -> voice_channel_id ペアを参照して処理
 - メッセージフィルタ (bot無視、コマンド無視、長さ制限 等)
 - メッセージ毎に TTS 合成を要求し VoiceChannel で再生
 - 再生キュー（asyncio.Queue）で順次再生（同時再生は基本抑制）
 - 再接続 / エラーハンドリングの基本
 - !start / !stop で TTS ON/OFF
 - !status / !join / !leave を受け付ける
 - Webhook へ「TTSを始めます」「TTSを終了します」「エラーが発生しました」を送信
"""
from __future__ import annotations

import asyncio
import logging
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Dict, Optional

import aiohttp
import discord  # discord.py 2.x
from discord import Intents, Message, VoiceChannel, VoiceClient

from config_manager import ConfigManager

LOGGER = logging.getLogger(__name__)

# 実行フォルダの基準を固定
if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent

# discord.py のバージョン警告
try:
    from packaging.version import Version

    if Version(discord.__version__) < Version("2.7.0"):
        LOGGER.warning(
            "discord.py %s が検出されました。voice 関連の安定性のため 2.7.0 以上を推奨します。",
            discord.__version__,
        )
except Exception:
    pass

# Try to import external tts_engine and message_processor if present.
# tts_engine must provide: async def synthesize_wav(text, speaker_id, speed) -> Path
try:
    import tts_engine  # type: ignore
except Exception:
    tts_engine = None
    LOGGER.warning("tts_engine モジュールが見つかりません。合成部分は未実装です。")

try:
    import message_processor  # type: ignore
except Exception:
    message_processor = None
    LOGGER.info("message_processor モジュールが見つかりません。内蔵の簡易処理を使います。")


# -------------------------
# ヘルパー
# -------------------------
URL_RE = re.compile(r"https?://\S+", re.IGNORECASE)


def is_url(text: str) -> bool:
    return bool(URL_RE.search(text))


def contains_image_attachment(msg: Message) -> bool:
    for a in msg.attachments:
        if a.filename:
            ext = a.filename.lower().split(".")[-1]
            if ext in ("png", "jpg", "jpeg", "gif", "webp", "bmp"):
                return True
    return False


def shutil_which(name: str) -> Optional[str]:
    """簡易 shutil.which"""
    try:
        return shutil.which(name)
    except Exception:
        return None


# -------------------------
# Queue アイテム定義
# -------------------------
class QueueItem:
    def __init__(
        self,
        guild_id: int,
        text_channel_id: int,
        voice_channel_id: int,
        user_id: int,
        content: str,
        original_message: Message,
    ):
        self.guild_id = guild_id
        self.text_channel_id = text_channel_id
        self.voice_channel_id = voice_channel_id
        self.user_id = user_id
        self.content = content
        self.original_message = original_message


# -------------------------
# DiscordService
# -------------------------
class DiscordService:
    def __init__(self, config: ConfigManager):
        """
        config: ConfigManager インスタンス
        """
        self.config = config
        self.client: Optional[discord.Client] = None
        self._client_task: Optional[asyncio.Task] = None
        self._queue: Optional[asyncio.Queue[QueueItem]] = None
        self._consumer_task: Optional[asyncio.Task] = None
        self._running = False

        # exe と同じディレクトリを基準にする
        self.base_dir = SCRIPT_DIR

        # 状態管理
        self._voice_clients: Dict[int, VoiceClient] = {}  # guild_id -> VoiceClient

        # 接続中の多重実行を防ぐ
        self._connect_locks: Dict[int, asyncio.Lock] = {}

        # 初回 preconnect の重複防止
        self._preconnect_done = False

        # TTS ON/OFF（guild ごと）
        self._tts_enabled: Dict[int, bool] = {}

        # 同一ユーザー連投カウント（簡易）
        self._last_user_id: Optional[int] = None
        self._same_user_count: int = 0

        # Webhook
        self._http_session: Optional[aiohttp.ClientSession] = None
        self._webhook_url: str = str(self.config.get("webhook_url", "") or "").strip()

        # config による初期値
        qcfg = self.config.get_queue_config()
        max_size = int(qcfg.get("max_size", 50))
        self._queue = asyncio.Queue(maxsize=max_size)

        # concurrency lock for playing to avoid races
        self._play_lock = asyncio.Lock()

    # ---------------------------------
    # Public API
    # ---------------------------------
    async def start(self) -> None:
        """
        Discord client を生成し start する。
        非同期に呼ぶこと（AppController などから await される）。
        """
        if self._running:
            LOGGER.info("DiscordService: 既に起動済み")
            return

        token = self.config.get_bot_token()
        if not token:
            raise RuntimeError("Bot token が設定されていません (config.json)")

        if self._http_session is None or self._http_session.closed:
            self._http_session = aiohttp.ClientSession()

        # Intents: メッセージ読み取りに必要な設定
        intents = Intents.default()
        intents.message_content = True
        intents.messages = True
        intents.guilds = True
        intents.voice_states = True

        service = self  # closure

        # Internal client class with handlers
        class _Client(discord.Client):
            async def on_ready(self_inner):
                LOGGER.info("Discord client on_ready: %s", self_inner.user)

                # preconnect は 1 セッションで 1 回だけ
                if service._preconnect_done:
                    LOGGER.info("事前接続は既に実行済みのためスキップします")
                    return

                service._preconnect_done = True

                try:
                    await service._preconnect_all()
                except Exception:
                    LOGGER.exception("on_ready: preconnect で例外が発生しました")
                    await service._webhook_notify_error("on_ready: preconnect で例外が発生しました")

            async def on_message(self_inner, message: Message):
                await service._on_message(message)

            async def on_error(self_inner, event_method, *args, **kwargs):
                LOGGER.exception("Discord client error in %s", event_method)
                await service._webhook_notify_error(f"Discord client error in {event_method}")

            async def on_voice_state_update(self_inner, member, before, after):
                """
                Bot自身の切断・移動も含めて、VC状態の後始末を行う。
                """
                try:
                    if not self_inner.user:
                        return

                    is_self_bot = member.id == self_inner.user.id

                    # bot以外のメンバー更新は、退出判定だけ見る
                    if not is_self_bot and member.bot:
                        return

                    # bot自身の状態変化は、辞書の掃除を優先
                    if is_self_bot:
                        for guild_id, vc in list(service._voice_clients.items()):
                            try:
                                live_vc = service._find_live_voice_client(guild_id)
                                if live_vc is None:
                                    service._voice_clients.pop(guild_id, None)
                                    continue

                                if not live_vc.is_connected():
                                    service._voice_clients.pop(guild_id, None)
                                    continue

                                if live_vc.channel is None:
                                    service._voice_clients.pop(guild_id, None)
                                    continue

                                service._voice_clients[guild_id] = live_vc
                            except Exception:
                                LOGGER.exception("bot自身の voice state 後始末で例外")
                                service._voice_clients.pop(guild_id, None)
                        return

                    # 一般メンバーの変化: その guild の VC が空なら退出
                    for guild_id, vc in list(service._voice_clients.items()):
                        if vc is None:
                            service._voice_clients.pop(guild_id, None)
                            continue

                        if not vc.is_connected():
                            service._voice_clients.pop(guild_id, None)
                            continue

                        channel = vc.channel
                        if channel is None:
                            service._voice_clients.pop(guild_id, None)
                            continue

                        non_bot_members = [m for m in channel.members if not m.bot]

                        if len(non_bot_members) == 0:
                            LOGGER.info("VCに誰もいないため退出します: guild=%s", guild_id)
                            try:
                                await vc.disconnect(force=True)
                            except Exception:
                                LOGGER.exception("自動退出失敗")
                                await service._webhook_notify_error(f"自動退出失敗 guild={guild_id}")
                            service._voice_clients.pop(guild_id, None)
                except Exception:
                    LOGGER.exception("on_voice_state_update で例外が発生しました")
                    await service._webhook_notify_error("on_voice_state_update で例外が発生しました")

        self.client = _Client(intents=intents)

        self._client_task = asyncio.create_task(self._run_client(token))
        self._consumer_task = asyncio.create_task(self._consumer_loop())

        self._running = True
        LOGGER.info("DiscordService: 起動タスクを開始しました")
        await self._webhook_notify_started()

    async def stop(self) -> None:
        """
        停止処理（非同期呼び出し）
        """
        if not self._running:
            LOGGER.info("DiscordService: 既に停止済み")
            return

        LOGGER.info("DiscordService: 停止開始")

        if self._consumer_task:
            self._consumer_task.cancel()
            try:
                await self._consumer_task
            except asyncio.CancelledError:
                LOGGER.debug("consumer task cancelled")
            except Exception:
                LOGGER.exception("consumer task 停止時に例外")

        for guild_id, vc in list(self._voice_clients.items()):
            try:
                if vc and vc.is_connected():
                    await vc.disconnect(force=True)
            except Exception:
                LOGGER.exception("VoiceClient disconnect error for guild %s", guild_id)
        self._voice_clients.clear()

        if self.client:
            try:
                await self.client.close()
            except Exception:
                LOGGER.exception("Discord client close error")

        if self._client_task:
            try:
                await self._client_task
            except asyncio.CancelledError:
                LOGGER.debug("client task cancelled")
            except Exception:
                LOGGER.debug("client task finished/errored")

        if self._http_session and not self._http_session.closed:
            try:
                await self._http_session.close()
            except Exception:
                LOGGER.exception("HTTP session close error")

        self._running = False
        LOGGER.info("DiscordService: 停止完了")
        await self._webhook_notify_stopped()

    def stop_sync(self) -> None:
        if self._running:
            try:
                loop = asyncio.get_running_loop()
                loop.create_task(self.stop())
            except RuntimeError:
                LOGGER.debug("stop_sync: running loop がありません")
            except Exception:
                LOGGER.debug("stop_sync: asyncio task の作成に失敗しました")

    # ---------------------------------
    # Webhook
    # ---------------------------------
    async def _send_webhook(self, content: str) -> None:
        if not self._webhook_url:
            return
        if self._http_session is None or self._http_session.closed:
            self._http_session = aiohttp.ClientSession()

        try:
            async with self._http_session.post(
                self._webhook_url,
                json={"content": content},
                timeout=aiohttp.ClientTimeout(total=10),
            ) as resp:
                if resp.status >= 400:
                    body = await resp.text()
                    LOGGER.warning("Webhook送信失敗 status=%s body=%s", resp.status, body[:300])
        except Exception:
            LOGGER.exception("Webhook送信で例外が発生しました")

    async def _webhook_notify_started(self) -> None:
        await self._send_webhook("TTSを開始しました")

    async def _webhook_notify_stopped(self) -> None:
        await self._send_webhook("TTSを終了します")

    async def _webhook_notify_error(self, detail: str = "") -> None:
        msg = "エラーが発生しました"
        if detail:
            msg += f" : {detail}"
        await self._send_webhook(msg)

    # ---------------------------------
    # 内部: discord クライアント起動
    # ---------------------------------
    async def _run_client(self, token: str) -> None:
        try:
            await self.client.start(token)  # type: ignore
        except asyncio.CancelledError:
            LOGGER.info("Discord client task cancelled")
        except Exception:
            LOGGER.exception("Discord client start failed")
            await self._webhook_notify_error("Discord client start failed")

    # ---------------------------------
    # 内部: command / state
    # ---------------------------------
    def _get_tts_enabled(self, guild_id: int) -> bool:
        return self._tts_enabled.get(guild_id, True)

    def _set_tts_enabled(self, guild_id: int, enabled: bool) -> None:
        self._tts_enabled[guild_id] = enabled

    def _is_control_command(self, content: str) -> Optional[str]:
        """
        !start / !stop / !status / !join / !leave を判定して返す。
        """
        text = content.strip().lower()
        if text == "!start":
            return "start"
        if text == "!stop":
            return "stop"
        if text == "!status":
            return "status"
        if text == "!join":
            return "join"
        if text == "!leave":
            return "leave"
        return None

    async def _handle_control_command(self, message: Message, command: str) -> None:
        guild = message.guild
        guild_id = int(guild.id) if guild else 0

        if command == "start":
            self._set_tts_enabled(guild_id, True)
            LOGGER.info("TTS ON: guild=%s", guild_id)
            await self._send_webhook("TTSを開始しました")
            try:
                await message.channel.send("🔊 TTSを開始しました")
            except Exception:
                pass
            return

        if command == "stop":
            self._set_tts_enabled(guild_id, False)
            LOGGER.info("TTS OFF: guild=%s", guild_id)
            await self._send_webhook("TTSを終了します")
            try:
                await message.channel.send("🔇 TTSを停止しました")
            except Exception:
                pass
            return

        if command == "status":
            state = self._tts_enabled.get(guild_id, True)
            try:
                await message.channel.send(f"📊 状態: {'ON' if state else 'OFF'}")
            except Exception:
                pass
            return

        if command == "join":
            if guild is None:
                try:
                    await message.channel.send("⚠ サーバー内で実行してください")
                except Exception:
                    pass
                return

            voice = getattr(message.author, "voice", None)
            if voice and voice.channel:
                vc = await self._ensure_voice_client(guild_id, voice.channel)
                if vc:
                    try:
                        await message.channel.send("✅ VCに接続しました")
                    except Exception:
                        pass
                else:
                    try:
                        await message.channel.send("⚠ VCに接続できませんでした")
                    except Exception:
                        pass
            else:
                try:
                    await message.channel.send("⚠ VCに入ってください")
                except Exception:
                    pass
            return

        if command == "leave":
            await self._cleanup_voice_client(guild_id)
            try:
                await message.channel.send("👋 VCから切断しました")
            except Exception:
                pass
            return

    # ---------------------------------
    # 内部: voice client 管理
    # ---------------------------------
    def _get_connect_lock(self, guild_id: int) -> asyncio.Lock:
        lock = self._connect_locks.get(guild_id)
        if lock is None:
            lock = asyncio.Lock()
            self._connect_locks[guild_id] = lock
        return lock

    def _find_live_voice_client(self, guild_id: int) -> Optional[VoiceClient]:
        """
        self._voice_clients と discord.py 側の voice_clients から、
        guild_id に対応する生きている VoiceClient を探す。
        """
        if self.client is not None:
            try:
                for vc in getattr(self.client, "voice_clients", []):
                    try:
                        if vc.guild and vc.guild.id == guild_id:
                            return vc
                    except Exception:
                        continue
            except Exception:
                pass

        vc = self._voice_clients.get(guild_id)
        if vc is not None:
            return vc

        return None

    async def _ensure_voice_client(
        self,
        guild_id: int,
        voice_channel: VoiceChannel,
        timeout: float = 20.0,
    ) -> Optional[VoiceClient]:
        """
        guild_id に対して、voice_channel に接続済みの VoiceClient を返す。
        接続済みなら再利用し、別チャンネル/壊れた状態なら再接続する。
        """
        lock = self._get_connect_lock(guild_id)
        async with lock:
            existing = self._find_live_voice_client(guild_id)

            try:
                if existing is not None and existing.is_connected():
                    if existing.channel and existing.channel.id == voice_channel.id:
                        self._voice_clients[guild_id] = existing
                        return existing

                    try:
                        await existing.disconnect(force=True)
                    except Exception:
                        LOGGER.exception("既存VoiceClientの切断に失敗しました")
                    finally:
                        self._voice_clients.pop(guild_id, None)

                elif existing is not None:
                    self._voice_clients.pop(guild_id, None)
            except Exception:
                LOGGER.exception("既存VoiceClientの状態確認で例外")
                self._voice_clients.pop(guild_id, None)

            existing = self._find_live_voice_client(guild_id)
            if existing is not None and existing.is_connected():
                if existing.channel and existing.channel.id == voice_channel.id:
                    self._voice_clients[guild_id] = existing
                    return existing

                try:
                    await existing.disconnect(force=True)
                except Exception:
                    LOGGER.exception("接続先が違う既存VoiceClientの切断に失敗しました")
                finally:
                    self._voice_clients.pop(guild_id, None)

            try:
                LOGGER.info("VoiceChannel に接続します: %s", voice_channel.id)
                vc = await voice_channel.connect(timeout=timeout, reconnect=True)
                self._voice_clients[guild_id] = vc
                return vc
            except Exception:
                LOGGER.exception("VoiceChannel への接続に失敗しました")
                self._voice_clients.pop(guild_id, None)

                try:
                    if existing is not None and existing.is_connected():
                        await existing.disconnect(force=True)
                except Exception:
                    pass

                await self._webhook_notify_error(f"VoiceChannel への接続に失敗しました guild={guild_id}")
                return None

    async def _cleanup_voice_client(self, guild_id: int) -> None:
        """
        guild_id の VoiceClient を安全に掃除する。
        """
        vc = self._voice_clients.get(guild_id)
        if vc is None:
            self._voice_clients.pop(guild_id, None)
            return

        try:
            if vc.is_connected():
                await vc.disconnect(force=True)
        except Exception:
            LOGGER.exception("VoiceClient cleanup failed for guild %s", guild_id)
        finally:
            self._voice_clients.pop(guild_id, None)

    async def _preconnect_all(self) -> None:
        """
        on_ready 時に呼ぶ事前接続処理。
        """
        if not self.client:
            return

        pairs = self.config.get_channel_pairs()
        for p in pairs:
            try:
                if not p.get("enabled", True):
                    continue

                vid = p.get("voice_channel_id", "")
                if not vid:
                    continue

                try:
                    vid_int = int(vid)
                except Exception:
                    LOGGER.warning("voice_channel_id が整数でないためスキップ: %s", vid)
                    continue

                ch = self.client.get_channel(vid_int)
                if ch is None:
                    try:
                        ch = await self.client.fetch_channel(vid_int)  # type: ignore
                    except Exception:
                        LOGGER.exception("Voice channel fetch failed: %s", vid_int)
                        ch = None

                if ch is None:
                    LOGGER.warning("Voice channel が見つかりません: %s", vid_int)
                    continue

                if not isinstance(ch, VoiceChannel):
                    LOGGER.warning("指定チャネルは VoiceChannel ではありません: %s", vid_int)
                    continue

                LOGGER.info("事前接続: %s (guild=%s)", vid_int, getattr(ch.guild, "id", None))
                vc = await self._ensure_voice_client(ch.guild.id, ch, timeout=20.0)
                if vc is not None:
                    LOGGER.info("事前接続成功: guild=%s vc=%s", ch.guild.id, vid_int)
                else:
                    LOGGER.warning("事前接続に失敗しました: %s", vid_int)
            except Exception:
                LOGGER.exception("事前接続に失敗しました")
                await self._webhook_notify_error("事前接続に失敗しました")

    # ---------------------------------
    # 内部: メッセージ受信処理
    # ---------------------------------
    async def _on_message(self, message: Message) -> None:
        try:
            if not self.client:
                LOGGER.debug("client not ready")
                return

            if message.author == self.client.user:
                return

            cfg_filters = self.config.get_filters()
            if cfg_filters.get("ignore_bots", True) and message.author.bot:
                return

            if message.guild is None:
                return

            guild_id = int(message.guild.id)

            # !start / !stop / !status / !join / !leave を最優先で処理
            cmd = self._is_control_command(message.content or "")
            if cmd is not None:
                await self._handle_control_command(message, cmd)
                return

            # TTS OFF なら通常メッセージを無視
            if not self._get_tts_enabled(guild_id):
                LOGGER.debug("TTSがOFFのため無視: guild=%s", guild_id)
                return

            # find channel pair for this text channel
            text_ch_id_str = str(message.channel.id)
            channel_pairs = self.config.get_channel_pairs()
            pair = None
            for p in channel_pairs:
                if str(p.get("text_channel_id")) == text_ch_id_str and p.get("enabled", True):
                    pair = p
                    break
            if pair is None:
                LOGGER.debug("テキストチャンネル %s は設定にありません", message.channel.id)
                return

            # ignore command messages? （!start 等は上で処理済み）
            if cfg_filters.get("ignore_commands", True):
                prefix = str(cfg_filters.get("command_prefix", "!"))
                if message.content.strip().startswith(prefix):
                    LOGGER.debug("コマンドと判定して無視: %s", message.content)
                    return

            # ignore empty
            if cfg_filters.get("ignore_empty", True) and (not message.content.strip()) and (not message.attachments):
                LOGGER.debug("空メッセージを無視")
                return

            # create a processed text (use message_processor if available)
            processed_text: Optional[str] = None
            if message_processor and hasattr(message_processor, "process"):
                try:
                    res = message_processor.process(message)  # type: ignore
                    if asyncio.iscoroutine(res):
                        processed_text = await res  # pragma: no cover
                    else:
                        processed_text = res
                except Exception:
                    LOGGER.exception("外部 message_processor 呼び出し失敗。内部処理を使います。")

            if processed_text is None:
                processed_text = self._simple_process_message(message)

            processed_text = str(processed_text)

            # final length check
            max_len = int(cfg_filters.get("max_length", 200))
            if len(processed_text) > max_len:
                processed_text = processed_text[:max_len] + "（省略）"

            # enqueue
            voice_channel_id = int(pair.get("voice_channel_id"))
            guild_id = int(pair.get("guild_id") or (message.guild.id if message.guild else 0))
            qi = QueueItem(
                guild_id=guild_id,
                text_channel_id=message.channel.id,
                voice_channel_id=voice_channel_id,
                user_id=message.author.id,
                content=processed_text,
                original_message=message,
            )

            try:
                self._queue.put_nowait(qi)
                LOGGER.debug(
                    "メッセージをキューに追加しました: user=%s chan=%s",
                    message.author.id,
                    message.channel.id,
                )
            except asyncio.QueueFull:
                qcfg = self.config.get_queue_config()
                if qcfg.get("drop_old_when_full", True):
                    try:
                        _ = self._queue.get_nowait()
                        self._queue.put_nowait(qi)
                        LOGGER.warning("キューが満杯のため古いメッセージを破棄しました")
                    except Exception:
                        LOGGER.exception("キューの置換に失敗しました")
                        await self._webhook_notify_error("キューの置換に失敗しました")
                else:
                    LOGGER.warning("キューが満杯のためメッセージを無視しました")

        except Exception:
            LOGGER.exception("on_message で例外が発生しました")
            await self._webhook_notify_error("on_message で例外が発生しました")

    # ---------------------------------
    # 内蔵メッセージ整形（置換・URL/画像処理）
    # ---------------------------------
    def _simple_process_message(self, message: Message) -> str:
        text = message.content or ""

        if contains_image_attachment(message):
            if text.strip():
                text = text + "（画像）"
            else:
                text = "画像"

        text = URL_RE.sub("URL", text)

        rules = self.config.get_replace_rules()
        for k in sorted(rules.keys(), key=lambda x: -len(x)):
            v = rules[k]
            try:
                text = text.replace(k, v)
            except Exception:
                continue

        return text.strip() or "（空メッセージ）"

    # ---------------------------------
    # Consumer: キューを逐次処理して再生
    # ---------------------------------
    async def _consumer_loop(self) -> None:
        LOGGER.info("consumer loop を開始します")
        try:
            while True:
                qi: QueueItem = await self._queue.get()
                try:
                    await self._handle_queue_item(qi)
                except Exception:
                    LOGGER.exception("queue item の処理で例外が発生しました")
                    await self._webhook_notify_error("queue item の処理で例外が発生しました")
                finally:
                    try:
                        self._queue.task_done()
                    except Exception:
                        pass
        except asyncio.CancelledError:
            LOGGER.info("consumer loop がキャンセルされました")
        except Exception:
            LOGGER.exception("consumer loop で例外が発生しました")
            await self._webhook_notify_error("consumer loop で例外が発生しました")

    async def _handle_queue_item(self, qi: QueueItem) -> None:
        LOGGER.info(
            "処理開始: guild=%s text_chan=%s voice_chan=%s user=%s",
            qi.guild_id,
            qi.text_channel_id,
            qi.voice_channel_id,
            qi.user_id,
        )

        # 1) speaker_id
        try:
            speaker_id = self.config.get_member_voice(str(qi.user_id))
        except Exception:
            speaker_id = self.config.get("default_speaker", 3)

        # 2) 合成
        speed = float(self.config.get("speed", 1.0))
        wav_path: Optional[Path] = None
        if tts_engine and hasattr(tts_engine, "synthesize_wav"):
            try:
                wav_path = await tts_engine.synthesize_wav(qi.content, speaker_id, speed)
                wav_path = Path(str(wav_path))
            except Exception:
                LOGGER.exception("tts_engine による合成失敗")
                await self._webhook_notify_error("tts_engine による合成失敗")
                wav_path = None
        else:
            LOGGER.error("tts_engine が利用できません。合成できません。")
            await self._webhook_notify_error("tts_engine が利用できません")
            return

        if wav_path is None or not wav_path.exists():
            LOGGER.error("合成された wav が見つかりません: %s", wav_path)
            await self._webhook_notify_error("合成された wav が見つかりません")
            return

        # 3) 接続
        try:
            if not self.client:
                LOGGER.error("discord client が未初期化")
                await self._webhook_notify_error("discord client が未初期化")
                return

            voice_channel = self.client.get_channel(qi.voice_channel_id)
            if voice_channel is None:
                try:
                    voice_channel = await self.client.fetch_channel(qi.voice_channel_id)  # type: ignore
                except Exception:
                    LOGGER.exception("voice channel の取得に失敗しました: %s", qi.voice_channel_id)
                    voice_channel = None

            if not isinstance(voice_channel, VoiceChannel):
                LOGGER.error("voice channel が VoiceChannel ではありません: %s", qi.voice_channel_id)
                await self._webhook_notify_error("voice channel が VoiceChannel ではありません")
                return

            vc = await self._ensure_voice_client(qi.guild_id, voice_channel, timeout=20.0)
            if vc is None:
                return

            # 4) 再生
            async with self._play_lock:
                if not vc.is_connected():
                    vc = await self._ensure_voice_client(qi.guild_id, voice_channel, timeout=20.0)
                    if vc is None:
                        return
                await self._play_wav_on_voice_client(vc, wav_path)

        except Exception:
            LOGGER.exception("queue item の再生処理で例外が発生しました")
            await self._webhook_notify_error("queue item の再生処理で例外が発生しました")
        finally:
            try:
                if wav_path and wav_path.exists():
                    tmpdir = os.getenv("TMP") or os.getenv("TEMP") or "/tmp"
                    if str(wav_path).startswith(str(tmpdir)) or "voxcord_temp" in str(wav_path).lower():
                        try:
                            wav_path.unlink()
                        except Exception:
                            LOGGER.debug("一時wav削除に失敗: %s", wav_path)
            except Exception:
                LOGGER.exception("wav削除処理で例外")

    async def _play_wav_on_voice_client(self, vc: VoiceClient, wav_path: Path) -> None:
        ffmpeg_exe = shutil_which("ffmpeg") or str(self.base_dir / "FFmpeg" / "ffmpeg.exe")
        if not ffmpeg_exe or not Path(ffmpeg_exe).exists():
            LOGGER.error("ffmpeg が見つかりません。再生できません。期待場所: %s", ffmpeg_exe)
            await self._webhook_notify_error("ffmpeg が見つかりません")
            return

        try:
            loop = asyncio.get_running_loop()
            play_done = loop.create_future()

            def _after_play(error):
                if error:
                    LOGGER.exception("再生中にエラー: %s", error)
                    try:
                        loop.call_soon_threadsafe(
                            lambda: asyncio.create_task(self._webhook_notify_error("再生中にエラー"))
                        )
                    except Exception:
                        pass
                if not play_done.done():
                    loop.call_soon_threadsafe(play_done.set_result, True)

            if vc.is_playing():
                try:
                    vc.stop()
                except Exception:
                    LOGGER.debug("既存再生を停止できませんでした")

            source = discord.FFmpegPCMAudio(executable=str(ffmpeg_exe), source=str(wav_path))
            vc.play(source, after=_after_play)

            LOGGER.info("再生開始: %s", wav_path)
            try:
                await asyncio.wait_for(play_done, timeout=300.0)
            except asyncio.TimeoutError:
                LOGGER.warning("再生タイムアウト。停止します")
                try:
                    vc.stop()
                except Exception:
                    pass
                await self._webhook_notify_error("再生タイムアウト")
            await asyncio.sleep(0.1)
        except Exception:
            LOGGER.exception("wav 再生で例外が発生しました")
            await self._webhook_notify_error("wav 再生で例外が発生しました")