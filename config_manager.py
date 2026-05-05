"""
ConfigManager

- 保存先: AppData\\Local\\VoxCord\\config.json を強制使用
- atomic save, バリデーション, ユーティリティメソッド多数
- スレッドセーフ (threading.RLock)
- コールバック登録可: 設定が保存されるとコールバックが呼ばれる
- import は AppData 側の config.json に反映
- export は指定先へ出力。未指定なら Downloads に出力
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import threading
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

LOGGER = logging.getLogger(__name__)


def _get_appdata_root() -> Path:
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / "VoxCord"
    return Path.home() / "AppData" / "Local" / "VoxCord"


def _get_downloads_dir() -> Path:
    userprofile = os.getenv("USERPROFILE")
    if userprofile:
        return Path(userprofile) / "Downloads"
    return Path.home() / "Downloads"


class ConfigManager:
    def __init__(self, path: str | Path = "config.json"):
        """
        path は受け取るが、保存先は AppData\\Local\\VoxCord\\config.json に固定する。
        """
        self._lock = threading.RLock()
        self._data: Dict[str, Any] = {}
        self._listeners: List[Callable[[Dict[str, Any]], None]] = []

        self.appdata_root = _get_appdata_root()
        self.appdata_root.mkdir(parents=True, exist_ok=True)

        self.path = self._resolve_config_path(path)

        try:
            self._load_or_create()
        except Exception:
            LOGGER.exception("初期化時にエラーが発生しました。デフォルトで再生成します。")
            with self._lock:
                self._data = self._default_config()
                try:
                    self._atomic_save(self._data)
                except Exception:
                    LOGGER.exception("デフォルト設定の保存に失敗しました。")

    def _resolve_config_path(self, path: str | Path) -> Path:
        """
        どんな path が来ても AppData\\Local\\VoxCord\\config.json に寄せる。
        ただし、拡張子やファイル名だけは使う。
        """
        try:
            p = Path(path)
            filename = p.name if p.name else "config.json"
            if not filename.lower().endswith(".json"):
                filename = "config.json"
        except Exception:
            filename = "config.json"
        return self.appdata_root / filename

    # -------------------------
    # デフォルト設定
    # -------------------------
    def _default_config(self) -> Dict[str, Any]:
        return {
            "bot_token": "",
            "default_speaker": 3,
            "speed": 1.0,
            "channel_pairs": [
                # {
                #   "guild_id": "",
                #   "text_channel_id": "",
                #   "voice_channel_id": "",
                #   "enabled": True
                # }
            ],
            "member_voice_map": {
                # "user_id": {"speaker_id": 8, "enabled": True}
            },
            "replace_rules": {
                # "w": "わら"
            },
            "filters": {
                "ignore_bots": True,
                "ignore_commands": True,
                "command_prefix": "!",
                "max_length": 200,
                "ignore_empty": True,
            },
            "queue": {
                "max_size": 50,
                "drop_old_when_full": True,
            },
        }

    # -------------------------
    # ロード / セーブ
    # -------------------------
    def _load_or_create(self):
        """ファイルがあれば読み込み、なければデフォルトを作成して保存する"""
        with self._lock:
            if not self.path.exists():
                LOGGER.info("config.json が見つかりません。デフォルトを作成します。")
                self._data = self._default_config()
                self._atomic_save(self._data)
                return

            try:
                raw = self.path.read_text(encoding="utf-8")
                loaded = json.loads(raw)
                if not isinstance(loaded, dict):
                    raise ValueError("config.json のルートがオブジェクトではありません")
                merged = self._merge_with_defaults(loaded)
                self._data = merged
                self._atomic_save(self._data)
            except Exception:
                ts = int(time.time())
                bak = self.path.parent / f"{self.path.name}.broken.{ts}.bak"
                try:
                    shutil.move(str(self.path), str(bak))
                    LOGGER.warning("壊れた設定ファイルをバックアップしました: %s", bak)
                except Exception:
                    LOGGER.exception("壊れた設定ファイルのバックアップに失敗しました")
                self._data = self._default_config()
                self._atomic_save(self._data)

    def _merge_with_defaults(self, loaded: Dict[str, Any]) -> Dict[str, Any]:
        defaults = self._default_config()
        merged = dict(defaults)
        for k, v in loaded.items():
            if k in defaults:
                if isinstance(defaults[k], dict) and isinstance(v, dict):
                    nested = dict(defaults[k])
                    nested.update(v)
                    merged[k] = nested
                else:
                    merged[k] = v
            else:
                merged[k] = v
        return merged

    def _atomic_save(self, data: Dict[str, Any]) -> None:
        """一時ファイル経由で原子的に保存"""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = self.path.with_suffix(self.path.suffix + ".tmp")
        text = json.dumps(data, ensure_ascii=False, indent=2)
        tmp_path.write_text(text, encoding="utf-8")
        tmp_path.replace(self.path)
        LOGGER.debug("config を保存しました: %s", str(self.path))

    def save(self) -> None:
        """現在の設定をファイルに保存し、リスナに通知する"""
        with self._lock:
            self._atomic_save(self._data)
            snapshot = dict(self._data)

        for cb in list(self._listeners):
            try:
                cb(snapshot)
            except Exception:
                LOGGER.exception("設定変更リスナで例外が発生しました")

    def reload(self) -> None:
        """ファイルから再ロードして内部データを更新"""
        with self._lock:
            if not self.path.exists():
                LOGGER.warning("reload: config ファイルが存在しません")
                return
            raw = self.path.read_text(encoding="utf-8")
            loaded = json.loads(raw)
            if not isinstance(loaded, dict):
                raise ValueError("config.json のルートがオブジェクトではありません")
            merged = self._merge_with_defaults(loaded)
            self._data = merged

        for cb in list(self._listeners):
            try:
                cb(dict(self._data))
            except Exception:
                LOGGER.exception("設定変更リスナで例外が発生しました")

    # -------------------------
    # 基本的な get/set
    # -------------------------
    def to_dict(self) -> Dict[str, Any]:
        with self._lock:
            return dict(self._data)

    def get(self, key: str, default: Any = None) -> Any:
        with self._lock:
            return self._data.get(key, default)

    def set(self, key: str, value: Any, save: bool = True) -> None:
        with self._lock:
            self._data[key] = value
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    # -------------------------
    # ボット・チャネル関連ユーティリティ
    # -------------------------
    def get_bot_token(self) -> str:
        return str(self.get("bot_token", "")) or ""

    def set_bot_token(self, token: str, save: bool = True) -> None:
        self.set("bot_token", token, save=save)

    def get_channel_pairs(self) -> List[Dict[str, Any]]:
        with self._lock:
            return list(self._data.get("channel_pairs", []))

    def add_channel_pair(
        self,
        guild_id: str,
        text_channel_id: str,
        voice_channel_id: str,
        enabled: bool = True,
        save: bool = True,
    ) -> None:
        with self._lock:
            pairs: List[Dict[str, Any]] = self._data.setdefault("channel_pairs", [])
            entry = {
                "guild_id": str(guild_id),
                "text_channel_id": str(text_channel_id),
                "voice_channel_id": str(voice_channel_id),
                "enabled": bool(enabled),
            }
            pairs.append(entry)
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    def remove_channel_pair(
        self,
        text_channel_id: Optional[str] = None,
        voice_channel_id: Optional[str] = None,
        save: bool = True,
    ) -> int:
        if text_channel_id is None and voice_channel_id is None:
            return 0

        with self._lock:
            pairs: List[Dict[str, Any]] = self._data.get("channel_pairs", [])
            before = len(pairs)
            pairs = [
                p for p in pairs
                if not (
                    (text_channel_id and p.get("text_channel_id") == str(text_channel_id)) or
                    (voice_channel_id and p.get("voice_channel_id") == str(voice_channel_id))
                )
            ]
            self._data["channel_pairs"] = pairs
            after = len(pairs)
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")
        return before - after

    def set_channel_pair_enabled(self, text_channel_id: str, enabled: bool, save: bool = True) -> bool:
        with self._lock:
            changed = False
            for p in self._data.get("channel_pairs", []):
                if p.get("text_channel_id") == str(text_channel_id):
                    p["enabled"] = bool(enabled)
                    changed = True
            if changed and save:
                self._atomic_save(self._data)

        if changed and save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")
        return changed

    # -------------------------
    # メンバー話者マップ
    # -------------------------
    def get_member_voice(self, user_id: str) -> Optional[int]:
        with self._lock:
            m = self._data.get("member_voice_map", {}).get(str(user_id))
            if not m or not m.get("enabled", True):
                return self._data.get("default_speaker")
            return int(m.get("speaker_id", self._data.get("default_speaker")))

    def set_member_voice(self, user_id: str, speaker_id: int, enabled: bool = True, save: bool = True) -> None:
        with self._lock:
            mv = self._data.setdefault("member_voice_map", {})
            mv[str(user_id)] = {"speaker_id": int(speaker_id), "enabled": bool(enabled)}
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    def remove_member_voice(self, user_id: str, save: bool = True) -> bool:
        with self._lock:
            mv: Dict[str, Any] = self._data.get("member_voice_map", {})
            if str(user_id) in mv:
                del mv[str(user_id)]
                if save:
                    self._atomic_save(self._data)
                if save:
                    for cb in list(self._listeners):
                        try:
                            cb(dict(self._data))
                        except Exception:
                            LOGGER.exception("設定変更リスナで例外が発生しました")
                return True
            return False

    # -------------------------
    # 置換ルール
    # -------------------------
    def get_replace_rules(self) -> Dict[str, str]:
        with self._lock:
            return dict(self._data.get("replace_rules", {}))

    def add_replace_rule(self, key: str, value: str, save: bool = True) -> None:
        with self._lock:
            rr = self._data.setdefault("replace_rules", {})
            rr[str(key)] = str(value)
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    def remove_replace_rule(self, key: str, save: bool = True) -> bool:
        with self._lock:
            rr: Dict[str, str] = self._data.get("replace_rules", {})
            if str(key) in rr:
                del rr[str(key)]
                if save:
                    self._atomic_save(self._data)
                if save:
                    for cb in list(self._listeners):
                        try:
                            cb(dict(self._data))
                        except Exception:
                            LOGGER.exception("設定変更リスナで例外が発生しました")
                return True
            return False

    # -------------------------
    # フィルタ / キュー系
    # -------------------------
    def get_filters(self) -> Dict[str, Any]:
        with self._lock:
            return dict(self._data.get("filters", {}))

    def set_filter(self, key: str, value: Any, save: bool = True) -> None:
        with self._lock:
            f = self._data.setdefault("filters", {})
            f[str(key)] = value
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    def get_queue_config(self) -> Dict[str, Any]:
        with self._lock:
            return dict(self._data.get("queue", {}))

    def set_queue_config(self, cfg: Dict[str, Any], save: bool = True) -> None:
        with self._lock:
            self._data["queue"] = dict(cfg)
            if save:
                self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    # -------------------------
    # エクスポート/インポート/バックアップ
    # -------------------------
    def export_to(self, out_path: str | Path | None = None) -> Path:
        """
        設定をエクスポートする。
        out_path が None の場合は Downloads に保存する。
        out_path がディレクトリなら、その中に config.json を作る。
        """
        if out_path is None or str(out_path).strip() == "":
            out = _get_downloads_dir() / "VoxCord_config.json"
        else:
            out = Path(out_path)
            if out.exists() and out.is_dir():
                out = out / "VoxCord_config.json"
            elif not out.suffix:
                out = out / "VoxCord_config.json"

        out.parent.mkdir(parents=True, exist_ok=True)

        with self._lock:
            text = json.dumps(self._data, ensure_ascii=False, indent=2)

        out.write_text(text, encoding="utf-8")
        LOGGER.info("設定をエクスポートしました: %s", out)
        return out

    def import_from(self, in_path: str | Path, save_after: bool = True) -> None:
        """
        設定を AppData 側の config.json に取り込む。
        つまり、読み込んだ設定は self.path に保存される。
        """
        p = Path(in_path)
        if not p.exists():
            raise FileNotFoundError(f"import ファイルが見つかりません: {p}")

        raw = p.read_text(encoding="utf-8")
        loaded = json.loads(raw)
        if not isinstance(loaded, dict):
            raise ValueError("import ファイルの形式が不正です")

        merged = self._merge_with_defaults(loaded)

        with self._lock:
            self._data = merged
            if save_after:
                self._atomic_save(self._data)

        if save_after:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")

    def backup(self, suffix: Optional[str] = None) -> Path:
        """設定ファイルをバックアップし、そのパスを返す"""
        with self._lock:
            if not self.path.exists():
                raise FileNotFoundError("バックアップ元の設定ファイルが存在しません")
            ts = int(time.time())
            sfx = suffix or f"{ts}"
            bak = self.path.parent / f"{self.path.name}.bak.{sfx}"
            shutil.copy2(self.path, bak)
        LOGGER.info("設定をバックアップしました: %s", bak)
        return bak

    # -------------------------
    # Listener (observer)
    # -------------------------
    def add_listener(self, cb: Callable[[Dict[str, Any]], None]) -> None:
        if not callable(cb):
            raise ValueError("listener は callable である必要があります")
        with self._lock:
            self._listeners.append(cb)

    def remove_listener(self, cb: Callable[[Dict[str, Any]], None]) -> None:
        with self._lock:
            try:
                self._listeners.remove(cb)
            except ValueError:
                pass

    # -------------------------
    # バリデーション（簡易）
    # -------------------------
    def validate(self) -> bool:
        """現在の設定に矛盾がないかをチェックし、問題があれば修正（可能な範囲で）する"""
        with self._lock:
            changed = False
            data = self._data

            try:
                data["speed"] = float(data.get("speed", 1.0))
            except Exception:
                data["speed"] = 1.0
                changed = True

            try:
                data["default_speaker"] = int(data.get("default_speaker", 3))
            except Exception:
                data["default_speaker"] = 3
                changed = True

            cp = data.get("channel_pairs", [])
            if not isinstance(cp, list):
                data["channel_pairs"] = []
                changed = True
            else:
                cleaned = []
                for item in cp:
                    if not isinstance(item, dict):
                        continue
                    item2 = {
                        "guild_id": str(item.get("guild_id", "")),
                        "text_channel_id": str(item.get("text_channel_id", "")),
                        "voice_channel_id": str(item.get("voice_channel_id", "")),
                        "enabled": bool(item.get("enabled", True)),
                    }
                    cleaned.append(item2)
                if cleaned != cp:
                    data["channel_pairs"] = cleaned
                    changed = True

            mvm = data.get("member_voice_map", {})
            if not isinstance(mvm, dict):
                data["member_voice_map"] = {}
                changed = True
            else:
                cleaned = {}
                for k, v in mvm.items():
                    try:
                        speaker = int(v.get("speaker_id", data.get("default_speaker", 3)))
                        enabled = bool(v.get("enabled", True))
                        cleaned[str(k)] = {"speaker_id": speaker, "enabled": enabled}
                    except Exception:
                        continue
                data["member_voice_map"] = cleaned

            rr = data.get("replace_rules", {})
            if not isinstance(rr, dict):
                data["replace_rules"] = {}
                changed = True
            else:
                cleaned = {str(k): str(v) for k, v in rr.items()}
                data["replace_rules"] = cleaned

            f = data.get("filters", {})
            if not isinstance(f, dict):
                data["filters"] = self._default_config()["filters"]
                changed = True

            q = data.get("queue", {})
            if not isinstance(q, dict):
                data["queue"] = self._default_config()["queue"]
                changed = True
            else:
                try:
                    data["queue"]["max_size"] = int(q.get("max_size", 50))
                except Exception:
                    data["queue"]["max_size"] = 50
                    changed = True

            if changed:
                try:
                    self._atomic_save(data)
                except Exception:
                    LOGGER.exception("validate 中の保存に失敗しました")
            return True

    # -------------------------
    # ヘルパー / レガシー補助
    # -------------------------
    def ensure_default_member(self, user_id: str, save: bool = True) -> None:
        """未登録ユーザーを member_voice_map にデフォルトで追加する（自動追加用）"""
        with self._lock:
            mv = self._data.setdefault("member_voice_map", {})
            if str(user_id) not in mv:
                mv[str(user_id)] = {
                    "speaker_id": int(self._data.get("default_speaker", 3)),
                    "enabled": True,
                }
                if save:
                    self._atomic_save(self._data)

        if save:
            for cb in list(self._listeners):
                try:
                    cb(dict(self._data))
                except Exception:
                    LOGGER.exception("設定変更リスナで例外が発生しました")
