"""
DiscordService - Discord 接続 / メッセージ監視 / 再生キュー管理

主な機能:
 - Bot接続（非同期）
 - 指定された text_channel_id -> voice_channel_id ペアを参照して処理
 - メッセージフィルタ (bot無視、コマンド無視、長さ制限 等)
 - メッセージ毎に TTS 合成を要求し VoiceChannel で再生
 - 再生キュー（asyncio.Queue）で順次再生（同時再生は基本抑制）
 - !join / !leave / !status / !skip / !clear / !volume
 - URL は「URL」、画像添付は「画像」と読み上げ
 - ユーザーごとの話者割り当て対応
 - Webhook へ状態やエラーを送信
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

import discord

discord.opus.load_opus("opus.dll")

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

# 外部 TTS / message_processor
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


def contains_image_attachment(msg: Message) -> bool:
    for a in msg.attachments:
        if a.filename:
            ext = a.filename.lower().split(".")[-1]
            if ext in ("png", "jpg", "jpeg", "gif", "webp", "bmp"):
                return True
    return False


def shutil_which(name: str) -> Optional[str]:
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

        # TTS ON/OFF（guild ごと）: 初期は ON 扱い
        self._tts_enabled: Dict[int, bool] = {}

        # 音量（1〜5、コマンドで循環）
        self._volume_levels: Dict[int, int] = {}

        # Webhook
        self._http_session: Optional[aiohttp.ClientSession] = None
        self._webhook_url: str = str(self.config.get("webhook_url", "") or "").strip()

        # config による初期値
        qcfg = self.config.get_queue_config()
        max_size = int(qcfg.get("max_size", 50))
        self._queue = asyncio.Queue(maxsize=max_size)

        # 再生競合防止
        self._play_lock = asyncio.Lock()

    # ---------------------------------
    # Public API
    # ---------------------------------
    async def start(self) -> None:
        if self._running:
            LOGGER.info("DiscordService: 既に起動済み")
            return

        token = self.config.get_bot_token()
        if not token:
            raise RuntimeError("Bot token が設定されていません (config.json)")

        if self._http_session is None or self._http_session.closed:
            self._http_session = aiohttp.ClientSession()

        intents = Intents.default()
        intents.message_content = True
        intents.messages = True
        intents.guilds = True
        intents.voice_states = True

        service = self

        class _Client(discord.Client):
            async def on_ready(self_inner):
                LOGGER.info("Discord client on_ready: %s", self_inner.user)

            async def on_message(self_inner, message: Message):
                await service._on_message(message)

            async def on_error(self_inner, event_method, *args, **kwargs):
                LOGGER.exception("Discord client error in %s", event_method)
                await service._webhook_notify_error(f"Discord client error in {event_method}")

            async def on_voice_state_update(self_inner, member, before, after):
                """
                bot自身の voice_client 状態だけ内部辞書に反映する。
                自動接続・自動退出はしない。
                """
                try:
                    if not self_inner.user:
                        return

                    if member.id != self_inner.user.id:
                        return

                    for guild_id, vc in list(service._voice_clients.items()):
                        try:
                            live_vc = service._find_live_voice_client(guild_id)
                            if live_vc is None:
                                service._voice_clients.pop(guild_id, None)
                                continue
                            if not live_vc.is_connected():
                                service._voice_clients.pop(guild_id, None)
                                continue
                            service._voice_clients[guild_id] = live_vc
                        except Exception:
                            LOGGER.exception("bot自身の voice state 後始末で例外")
                            service._voice_clients.pop(guild_id, None)
                except Exception:
                    LOGGER.exception("on_voice_state_update で例外が発生しました")

        self.client = _Client(intents=intents)

        self._client_task = asyncio.create_task(self._run_client(token))
        self._consumer_task = asyncio.create_task(self._consumer_loop())

        self._running = True
        LOGGER.info("DiscordService: 起動タスクを開始しました")
        await self._webhook_notify_started()

    async def stop(self) -> None:
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
    # コマンド / 状態
    # ---------------------------------
    def _get_tts_enabled(self, guild_id: int) -> bool:
        return self._tts_enabled.get(guild_id, True)

    def _set_tts_enabled(self, guild_id: int, enabled: bool) -> None:
        self._tts_enabled[guild_id] = enabled

    def _get_volume_level(self, guild_id: int) -> int:
        value = int(self._volume_levels.get(guild_id, 3))
        if value < 1 or value > 5:
            value = 3
        return value

    def _cycle_volume_level(self, guild_id: int) -> int:
        current = self._get_volume_level(guild_id)
        next_value = current + 1 if current < 5 else 1
        self._volume_levels[guild_id] = next_value
        return next_value

    def _is_control_command(self, content: str) -> Optional[str]:
        text = content.strip().lower()
        if text == "!status":
            return "status"
        if text == "!join":
            return "join"
        if text == "!leave":
            return "leave"
        if text == "!skip":
            return "skip"
        if text == "!clear":
            return "clear"
        if text == "!volume":
            return "volume"
        return None

    async def _handle_control_command(self, message: Message, command: str) -> None:
        guild = message.guild
        guild_id = int(guild.id) if guild else 0

        if command == "status":
            vc = self._find_live_voice_client(guild_id)
            connected = bool(vc and vc.is_connected())
            vc_name = str(vc.channel.name) if vc and vc.channel else "未接続"
            queue_size = self._queue.qsize() if self._queue else 0
            volume = self._get_volume_level(guild_id)
            state = "接続中" if connected else "未接続"
            try:
                await message.channel.send(
                    f"📊 状態: {state} / VC: {vc_name} / キュー: {queue_size} / 音量: {volume}/5"
                )
            except Exception:
                pass
            return

        if command == "join":
            target = await self._resolve_target_voice_channel(message)
            if target is None:
                try:
                    await message.channel.send("⚠ 接続先のVoiceChannelが見つかりません")
                except Exception:
                    pass
                return

            vc = await self._ensure_voice_client(guild_id, target, timeout=20.0)
            if vc:
                try:
                    await message.channel.send(f"✅ VCに接続しました: {target.name}")
                except Exception:
                    pass
            else:
                try:
                    await message.channel.send("⚠ VCに接続できませんでした")
                except Exception:
                    pass
            return

        if command == "leave":
            cleared = await self._clear_queue()
            await self._cleanup_voice_client(guild_id)
            try:
                await message.channel.send(f"👋 切断しました / キューを {cleared} 件削除しました")
            except Exception:
                pass
            return

        if command == "skip":
            stopped = await self._skip_current_playback(guild_id)
            try:
                if stopped:
                    await message.channel.send("⏭ 現在の再生をスキップしました")
                else:
                    await message.channel.send("⚠ 再生中の音声はありません")
            except Exception:
                pass
            return

        if command == "clear":
            cleared = await self._clear_queue()
            try:
                await message.channel.send(f"🧹 キューを {cleared} 件削除しました")
            except Exception:
                pass
            return

        if command == "volume":
            new_volume = self._cycle_volume_level(guild_id)
            try:
                await message.channel.send(f"🔊 音量を {new_volume}/5 に変更しました")
            except Exception:
                pass
            return

    async def _resolve_target_voice_channel(self, message: Message) -> Optional[VoiceChannel]:
        """
        join 用の接続先を解決する。
        まず text_channel_id に一致する設定を探し、なければ同一 guild 内の enabled 1件を使う。
        """
        if message.guild is None:
            return None

        target_pair = None
        text_ch_id = str(message.channel.id)
        guild_id = str(message.guild.id)

        for p in self.config.get_channel_pairs():
            if not p.get("enabled", True):
                continue
            if str(p.get("text_channel_id")) == text_ch_id:
                target_pair = p
                break

        if target_pair is None:
            enabled_pairs = [
                p for p in self.config.get_channel_pairs()
                if p.get("enabled", True) and str(p.get("guild_id", "")) == guild_id
            ]
            if len(enabled_pairs) == 1:
                target_pair = enabled_pairs[0]

        if target_pair is None:
            return None

        try:
            voice_channel_id = int(target_pair.get("voice_channel_id"))
        except Exception:
            return None

        if self.client is None:
            return None

        ch = self.client.get_channel(voice_channel_id)
        if ch is None:
            try:
                ch = await self.client.fetch_channel(voice_channel_id)  # type: ignore
            except Exception:
                LOGGER.exception("Voice channel の取得に失敗しました: %s", voice_channel_id)
                ch = None

        if not isinstance(ch, VoiceChannel):
            return None

        return ch

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
        if self.client is not None:
            try:
                guild = self.client.get_guild(guild_id)
                if guild is not None:
                    live = getattr(guild, "voice_client", None)
                    if live is not None and live.is_connected():
                        return live
            except Exception:
                pass

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
        join コマンドでのみ使う。
        """
        lock = self._get_connect_lock(guild_id)
        async with lock:
            try:
                guild = voice_channel.guild
                existing = None

                if guild is not None:
                    existing = getattr(guild, "voice_client", None)

                if existing is None:
                    existing = self._find_live_voice_client(guild_id)

                if existing is not None:
                    try:
                        if existing.is_connected():
                            if existing.channel and existing.channel.id == voice_channel.id:
                                self._voice_clients[guild_id] = existing
                                return existing

                            try:
                                await existing.disconnect(force=True)
                            except Exception:
                                LOGGER.exception("既存VoiceClientの切断に失敗しました")
                            finally:
                                self._voice_clients.pop(guild_id, None)
                    except Exception:
                        LOGGER.exception("既存VoiceClientの状態確認で例外")
                        self._voice_clients.pop(guild_id, None)

                try:
                    live_after = None
                    if voice_channel.guild is not None:
                        live_after = getattr(voice_channel.guild, "voice_client", None)
                    if live_after is not None and live_after.is_connected():
                        if live_after.channel and live_after.channel.id == voice_channel.id:
                            self._voice_clients[guild_id] = live_after
                            return live_after
                        try:
                            await live_after.disconnect(force=True)
                        except Exception:
                            LOGGER.exception("接続先が違う既存VoiceClientの切断に失敗しました")
                        finally:
                            self._voice_clients.pop(guild_id, None)
                except Exception:
                    LOGGER.exception("再確認時のVoiceClient状態確認で例外")

                LOGGER.info("VoiceChannel に接続します: %s", voice_channel.id)

                try:
                    vc = await voice_channel.connect(timeout=timeout, reconnect=True)
                    self._voice_clients[guild_id] = vc
                    return vc
                except discord.ClientException as e:
                    if "Already connected to a voice channel" in str(e):
                        live = getattr(voice_channel.guild, "voice_client", None) if voice_channel.guild else None
                        if live is not None and live.is_connected():
                            self._voice_clients[guild_id] = live
                            return live
                    raise

            except Exception:
                LOGGER.exception("VoiceChannel への接続に失敗しました")
                self._voice_clients.pop(guild_id, None)

                try:
                    existing = self._find_live_voice_client(guild_id)
                    if existing is not None and existing.is_connected():
                        await existing.disconnect(force=True)
                except Exception:
                    pass

                await self._webhook_notify_error(f"VoiceChannel への接続に失敗しました guild={guild_id}")
                return None

    async def _cleanup_voice_client(self, guild_id: int) -> None:
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

    async def _skip_current_playback(self, guild_id: int) -> bool:
        vc = self._find_live_voice_client(guild_id)
        if vc is None or not vc.is_connected():
            return False
        try:
            if vc.is_playing():
                vc.stop()
                return True
            return False
        except Exception:
            LOGGER.exception("skip_current_playback failed guild=%s", guild_id)
            return False

    async def _clear_queue(self) -> int:
        if self._queue is None:
            return 0

        cleared = 0
        while True:
            try:
                self._queue.get_nowait()
                self._queue.task_done()
                cleared += 1
            except asyncio.QueueEmpty:
                break
            except Exception:
                break
        return cleared

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

            # コマンドは最優先
            cmd = self._is_control_command(message.content or "")
            if cmd is not None:
                await self._handle_control_command(message, cmd)
                return

            # 通常メッセージの無視条件
            if cfg_filters.get("ignore_commands", True):
                prefix = str(cfg_filters.get("command_prefix", "!"))
                if message.content.strip().startswith(prefix):
                    LOGGER.debug("コマンドと判定して無視: %s", message.content)
                    return

            if cfg_filters.get("ignore_empty", True) and (not message.content.strip()) and (not message.attachments):
                LOGGER.debug("空メッセージを無視")
                return

            # text->voice pair を確認
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

            # メッセージ整形
            processed_text: Optional[str] = None
            if message_processor and hasattr(message_processor, "process"):
                try:
                    res = message_processor.process(message)  # type: ignore
                    if asyncio.iscoroutine(res):
                        processed_text = await res
                    else:
                        processed_text = res
                except Exception:
                    LOGGER.exception("外部 message_processor 呼び出し失敗。内部処理を使います。")

            if processed_text is None:
                processed_text = self._simple_process_message(message)
            else:
                processed_text = self._normalize_post_process_text(str(processed_text), message)

            processed_text = str(processed_text)

            max_len = int(cfg_filters.get("max_length", 200))
            if len(processed_text) > max_len:
                processed_text = processed_text[:max_len] + "（省略）"

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

    def _normalize_post_process_text(self, text: str, message: Message) -> str:
        if contains_image_attachment(message):
            if text.strip():
                text = text + " 画像"
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
    # 内蔵メッセージ整形（置換・URL/画像処理）
    # ---------------------------------
    def _simple_process_message(self, message: Message) -> str:
        text = message.content or ""
        text = self._normalize_post_process_text(text, message)
        return text

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

        # speaker_id
        try:
            speaker_id = self.config.get_member_voice(str(qi.user_id))
        except Exception:
            speaker_id = self.config.get("default_speaker", 3)

        # 合成
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

        # 再生は「既に接続済みの VC がある場合のみ」
        try:
            if not self.client:
                LOGGER.error("discord client が未初期化")
                await self._webhook_notify_error("discord client が未初期化")
                return

            voice_client = self._find_live_voice_client(qi.guild_id)
            if voice_client is None or not voice_client.is_connected():
                LOGGER.warning("VoiceClient が接続されていないため再生をスキップします: guild=%s", qi.guild_id)
                return

            if not voice_client.channel or int(voice_client.channel.id) != int(qi.voice_channel_id):
                LOGGER.warning(
                    "VoiceClient の接続先が違うため再生をスキップします: guild=%s current=%s expected=%s",
                    qi.guild_id,
                    getattr(voice_client.channel, "id", None),
                    qi.voice_channel_id,
                )
                return

            async with self._play_lock:
                if not voice_client.is_connected():
                    LOGGER.warning("再生直前に VoiceClient が切断されていました: guild=%s", qi.guild_id)
                    return

                await self._play_wav_on_voice_client(voice_client, wav_path, qi.guild_id)

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

    async def _play_wav_on_voice_client(self, vc: VoiceClient, wav_path: Path, guild_id: int) -> None:
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
            volume = self._get_volume_level(guild_id) / 5.0
            source = discord.PCMVolumeTransformer(source, volume=volume)

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
