import re
import logging
import discord

LOGGER = logging.getLogger(__name__)


class MessageProcessor:
    def __init__(self, config_manager):
        self.config = config_manager

        # URL検出
        self.url_pattern = re.compile(r'https?://\S+')

        # w連続検出
        self.w_pattern = re.compile(r'(w|ｗ)+')

        # 絵文字検出（カスタム絵文字含む）
        self.emoji_pattern = re.compile(r'<a?:\w+:\d+>')

    # =========================
    # メイン処理
    # =========================
    def process(self, message: discord.Message) -> str:
        try:
            text = message.content or ""

            # 添付ファイル判定
            if message.attachments:
                text += " 画像"

            text = self._replace_urls(text)
            text = self._replace_custom_emoji(text)
            text = self._replace_mentions(text, message)
            text = self._replace_w(text)
            text = self._apply_custom_replacements(text)
            text = self._normalize_text(text)

            result = text.strip()

            LOGGER.debug("Processed message: %s -> %s", message.content, result)
            return result

        except Exception:
            LOGGER.exception("Message processing failed")
            return ""

    # =========================
    # URL → 「URL」
    # =========================
    def _replace_urls(self, text: str) -> str:
        return self.url_pattern.sub("URL", text)

    # =========================
    # カスタム絵文字 → 「絵文字」
    # =========================
    def _replace_custom_emoji(self, text: str) -> str:
        return self.emoji_pattern.sub("絵文字", text)

    # =========================
    # メンション → 名前
    # =========================
    def _replace_mentions(self, text: str, message: discord.Message) -> str:
        try:
            # ユーザー
            for user in message.mentions:
                text = text.replace(f"<@{user.id}>", user.display_name)
                text = text.replace(f"<@!{user.id}>", user.display_name)

            # ロール
            for role in message.role_mentions:
                text = text.replace(f"<@&{role.id}>", role.name)

            # チャンネル
            for channel in message.channel_mentions:
                text = text.replace(f"<#{channel.id}>", channel.name)

        except Exception:
            LOGGER.exception("Mention replace failed")

        return text

    # =========================
    # w連続 → わら
    # =========================
    def _replace_w(self, text: str) -> str:
        return self.w_pattern.sub("わら", text)

    # =========================
    # ユーザー定義置換
    # =========================
    def _apply_custom_replacements(self, text: str) -> str:
        try:
            replacements = self.config.get("replacements", {})
            for before, after in replacements.items():
                text = text.replace(before, after)
        except Exception:
            LOGGER.exception("Custom replacement failed")

        return text

    # =========================
    # テキスト整形
    # =========================
    def _normalize_text(self, text: str) -> str:
        text = text.replace("\n", "。")
        text = re.sub(r'。+', "。", text)
        text = re.sub(r'\s+', " ", text)
        return text
