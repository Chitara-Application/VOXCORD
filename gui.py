"""
VoxCord GUI (PySide6) - Modern Discord Themed
"""

from __future__ import annotations

import sys
import os
import json
from pathlib import Path
from typing import Optional, List, Dict

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QSlider, QTextEdit, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QLineEdit,
    QFileDialog, QMessageBox, QTabWidget, QHeaderView, QCheckBox, QSpinBox,
    QApplication,
)
from PySide6.QtGui import QPixmap, QIcon, QAction, QFont
from PySide6.QtCore import Qt, QTimer, QSize

# 実行元ディレクトリを固定
if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent

SPEAKERS_JSON = SCRIPT_DIR / "speakers.json"

# Try to import tts_engine to fetch available speakers (best-effort)
try:
    import tts_engine  # type: ignore
except Exception:
    tts_engine = None


class VoxCordGUI(QMainWindow):
    def __init__(self, config_manager, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.config = config_manager

        self.setWindowTitle("VoxCord - Discord TTS Bot")
        self.setMinimumSize(1000, 720)

        self._set_app_icon()

        # attributes expected by main.py
        self.start_button: QPushButton = None  # set later
        self.stop_button: QPushButton = None

        # internal
        self._log_timer = QTimer(self)
        self._log_timer.setInterval(500)
        self._log_timer.timeout.connect(lambda: None)

        self._build_ui()
        self._connect_signals()

        try:
            self.config.add_listener(self._on_config_changed)
        except Exception:
            pass

    # -------------------------
    # Helper: Load Icons safely
    # -------------------------
    def _icon(self, filename: str) -> QIcon:
        for base in (SCRIPT_DIR / "assets", SCRIPT_DIR):
            path = base / filename
            if path.exists():
                return QIcon(str(path))
        return QIcon()

    def _set_app_icon(self):
        icon = QIcon()
        for path in (
            SCRIPT_DIR / "logo.ico",
            SCRIPT_DIR / "assets" / "logo.ico",
            SCRIPT_DIR / "assets" / "logo.png",
            SCRIPT_DIR / "logo.png",
        ):
            if path.exists():
                icon = QIcon(str(path))
                break

        if not icon.isNull():
            self.setWindowIcon(icon)
            app = QApplication.instance()
            if app is not None:
                app.setWindowIcon(icon)

    def _load_speakers_from_json(self) -> List[Dict]:
        """
        speakers.json を読む。
        形式例:
        [
          {
            "name": "四国めたん",
            "speaker_uuid": "...",
            "styles": [
              {"name": "ノーマル", "id": 2, "type": "talk"}
            ]
          }
        ]
        """
        if not SPEAKERS_JSON.exists():
            return []

        try:
            raw = SPEAKERS_JSON.read_text(encoding="utf-8")
            data = json.loads(raw)
            if isinstance(data, list):
                return data
        except Exception:
            pass
        return []

    def _populate_speakers_into(self, combo: QComboBox):
        combo.clear()

        speakers = self._load_speakers_from_json()

        if speakers:
            for sp in speakers:
                base_name = str(sp.get("name", "speaker"))
                styles = sp.get("styles", [])
                if not isinstance(styles, list):
                    continue

                for style in styles:
                    try:
                        style_name = str(style.get("name", "unknown"))
                        sid = int(style.get("id"))
                        combo.addItem(f"{base_name} ({style_name}) [{sid}]", sid)
                    except Exception:
                        continue

            if combo.count() > 0:
                return

        try:
            fetched = []
            if tts_engine:
                if hasattr(tts_engine, "get_speakers"):
                    fetched = tts_engine.get_speakers()
                elif hasattr(tts_engine, "TTSEngine"):
                    engine = tts_engine.TTSEngine(str(SCRIPT_DIR))
                    fetched = engine.get_speakers()

            if fetched and isinstance(fetched, list):
                for s in fetched:
                    try:
                        name = s.get("name", str(s.get("id", "speaker")))
                        sid = int(s.get("id", s.get("speaker_id", 0)))
                        combo.addItem(f"{name} ({sid})", sid)
                    except Exception:
                        continue
                return
        except Exception:
            pass

        for i in range(0, 16):
            combo.addItem(f"speaker_{i} ({i})", i)

    # -------------------------
    # UI 構築
    # -------------------------
    def _build_ui(self):
        central = QWidget()
        central.setObjectName("CentralWidget")
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(24, 24, 24, 24)

        top_bar = QHBoxLayout()
        top_bar.setSpacing(16)
        main_layout.addLayout(top_bar)

        logo_label = QLabel()
        logo_path = SCRIPT_DIR / "assets" / "logo.png"
        if logo_path.exists():
            logo_pix = QPixmap(str(logo_path)).scaledToHeight(56, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pix)
        else:
            logo_label.setText("🎙️")
            logo_label.setStyleSheet("font-size: 36px;")
        top_bar.addWidget(logo_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        title_layout = QVBoxLayout()
        title_layout.setSpacing(2)
        title_label = QLabel("VOXCORD")
        title_label.setStyleSheet("font-weight: 900; font-size: 24px; color: #F2F3F5; letter-spacing: 2px;")
        subtitle_label = QLabel("Discord → Voicevox TTS Bot")
        subtitle_label.setStyleSheet("font-size: 13px; color: #B5BAC1; font-weight: bold;")
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        top_bar.addLayout(title_layout)

        top_bar.addStretch()

        self.start_button = QPushButton(" Start Bot")
        self.start_button.setObjectName("StartButton")
        self.start_button.setIcon(self._icon("start.png"))
        self.start_button.setIconSize(QSize(28, 28))
        self.start_button.setMinimumHeight(44)
        self.start_button.setCursor(Qt.PointingHandCursor)
        top_bar.addWidget(self.start_button)

        self.stop_button = QPushButton(" Stop")
        self.stop_button.setObjectName("StopButton")
        self.stop_button.setIcon(self._icon("stop.png"))
        self.stop_button.setIconSize(QSize(28, 28))
        self.stop_button.setMinimumHeight(44)
        self.stop_button.setCursor(Qt.PointingHandCursor)
        top_bar.addWidget(self.stop_button)

        self.status_label = QLabel("STOPPED")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFixedSize(120, 44)
        self._apply_status_style("STOPPED")
        top_bar.addWidget(self.status_label)

        self.tabs = QTabWidget()
        self.tabs.setIconSize(QSize(24, 24))
        main_layout.addWidget(self.tabs, stretch=1)

        self.tab_main = QWidget()
        self.tabs.addTab(self.tab_main, self._icon("logo.png"), " メイン")
        self._build_tab_main()

        self.tab_map = QWidget()
        self.tabs.addTab(self.tab_map, self._icon("mappings.png"), " 話者マッピング")
        self._build_tab_mapping()

        self.tab_cfg = QWidget()
        self.tabs.addTab(self.tab_cfg, self._icon("channel_and_bot_settings.png"), " Bot / チャンネル")
        self._build_tab_channels()

        self.tab_replace = QWidget()
        self.tabs.addTab(self.tab_replace, self._icon("replacement.png"), " 置換設定")
        self._build_tab_replace()

        self._apply_stylesheet()

    # -------------------------
    # Main tab
    # -------------------------
    def _build_tab_main(self):
        layout = QVBoxLayout(self.tab_main)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 24, 20, 20)

        controls = QHBoxLayout()
        layout.addLayout(controls)

        speed_label = QLabel("TTS Speed:")
        speed_label.setStyleSheet("font-weight: 800; color: #B5BAC1; font-size: 14px;")
        controls.addWidget(speed_label)

        self.speed_slider = QSlider(Qt.Horizontal)
        self.speed_slider.setMinimum(50)
        self.speed_slider.setMaximum(200)
        self.speed_slider.setValue(int(float(self.config.get("speed", 1.0)) * 100))
        self.speed_slider.setTickInterval(10)
        self.speed_slider.setFixedWidth(300)
        self.speed_slider.setCursor(Qt.PointingHandCursor)
        controls.addWidget(self.speed_slider)

        self.speed_value_label = QLabel(f"{self.config.get('speed', 1.0):.2f}x")
        self.speed_value_label.setStyleSheet("font-weight: 900; color: #F2F3F5; font-size: 14px;")
        self.speed_value_label.setFixedWidth(50)
        controls.addWidget(self.speed_value_label)
        controls.addStretch()

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText(">> System logs will appear here...")
        layout.addWidget(self.log_view, stretch=1)

        bottom = QHBoxLayout()
        layout.addLayout(bottom)

        self.export_btn = QPushButton(" 設定エクスポート")
        self.export_btn.setIcon(self._icon("export.png"))
        self.export_btn.setIconSize(QSize(20, 20))

        self.import_btn = QPushButton(" 設定インポート")
        self.import_btn.setIcon(self._icon("import.png"))
        self.import_btn.setIconSize(QSize(20, 20))

        bottom.addWidget(self.export_btn)
        bottom.addWidget(self.import_btn)
        bottom.addStretch()

    # -------------------------
    # Mapping tab
    # -------------------------
    def _build_tab_mapping(self):
        layout = QVBoxLayout(self.tab_map)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        top_row = QHBoxLayout()
        layout.addLayout(top_row)

        lbl = QLabel("デフォルト話者:")
        lbl.setStyleSheet("font-weight: 800; color: #B5BAC1;")
        top_row.addWidget(lbl)

        self.default_speaker_combo = QComboBox()
        self.default_speaker_combo.setEditable(False)
        self.default_speaker_combo.setMinimumWidth(250)
        self._populate_speakers_into(self.default_speaker_combo)
        self.default_speaker_combo.setCurrentIndex(0)
        top_row.addWidget(self.default_speaker_combo)
        top_row.addStretch()

        self.reload_members_btn = QPushButton(" リロード")
        self.reload_members_btn.setObjectName("SecondaryButton")
        self.reload_members_btn.setIcon(self._icon("reload.png"))
        self.reload_members_btn.setIconSize(QSize(20, 20))
        top_row.addWidget(self.reload_members_btn)

        self.map_table = QTableWidget(0, 5)
        self.map_table.setHorizontalHeaderLabels(["有効", "User ID", "表示名", "話者", "操作"])
        self.map_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.map_table.setColumnWidth(0, 70)
        self.map_table.setColumnWidth(1, 200)
        self.map_table.setColumnWidth(2, 220)
        layout.addWidget(self.map_table, stretch=1)

        bottom = QHBoxLayout()
        layout.addLayout(bottom)
        self.add_user_id_input = QLineEdit()
        self.add_user_id_input.setPlaceholderText("User ID を入力して追加...")

        self.add_user_btn = QPushButton(" 追加")
        self.add_user_btn.setObjectName("PrimaryButton")
        self.add_user_btn.setIcon(self._icon("add.png"))
        self.add_user_btn.setIconSize(QSize(20, 20))

        bottom.addWidget(self.add_user_id_input)
        bottom.addWidget(self.add_user_btn)

        self._refresh_map_table()

    # -------------------------
    # Channels / Bot tab
    # -------------------------
    def _build_tab_channels(self):
        layout = QVBoxLayout(self.tab_cfg)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 24, 20, 20)

        form_widget = QWidget()
        form_widget.setObjectName("FormContainer")
        grid = QGridLayout(form_widget)
        grid.setVerticalSpacing(16)
        grid.setHorizontalSpacing(16)
        grid.setContentsMargins(20, 20, 20, 20)
        layout.addWidget(form_widget)

        def make_label(text):
            lbl = QLabel(text)
            lbl.setStyleSheet("font-weight: 900; color: #87898C; text-transform: uppercase; font-size: 12px;")
            return lbl

        grid.addWidget(make_label("Bot Token"), 0, 0)
        self.token_input = QLineEdit()
        self.token_input.setEchoMode(QLineEdit.Password)
        self.token_input.setPlaceholderText("トークンを入力...")
        self.token_input.setText(self.config.get_bot_token())
        grid.addWidget(self.token_input, 0, 1, 1, 2)

        self.save_token_btn = QPushButton("保存")
        self.save_token_btn.setObjectName("PrimaryButton")
        grid.addWidget(self.save_token_btn, 0, 3)

        line = QWidget()
        line.setFixedHeight(1)
        line.setStyleSheet("background-color: #1E1F22;")
        grid.addWidget(line, 1, 0, 1, 4)

        grid.addWidget(make_label("SERVER ID"), 2, 0)
        self.guild_input = QLineEdit()
        self.guild_input.setPlaceholderText("000000000000000000")
        grid.addWidget(self.guild_input, 2, 1)

        grid.addWidget(make_label("Text Channel ID"), 2, 2)
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("000000000000000000")
        grid.addWidget(self.text_input, 2, 3)

        grid.addWidget(make_label("Voice Channel ID"), 3, 0)
        self.voice_input = QLineEdit()
        self.voice_input.setPlaceholderText("000000000000000000")
        grid.addWidget(self.voice_input, 3, 1)

        self.add_pair_btn = QPushButton(" ペアを追加")
        self.add_pair_btn.setObjectName("PrimaryButton")
        self.add_pair_btn.setIcon(self._icon("add.png"))
        self.add_pair_btn.setIconSize(QSize(20, 20))
        grid.addWidget(self.add_pair_btn, 3, 2, 1, 2)

        self.pair_table = QTableWidget(0, 5)
        self.pair_table.setHorizontalHeaderLabels(["有効", "Guild ID", "Text ID", "Voice ID", "操作"])
        self.pair_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.pair_table, stretch=1)

        self._refresh_channel_pairs()

    # -------------------------
    # Replace tab
    # -------------------------
    def _build_tab_replace(self):
        layout = QVBoxLayout(self.tab_replace)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        self.replace_table = QTableWidget(0, 3)
        self.replace_table.setHorizontalHeaderLabels(["元文字列", "読み", "操作"])
        self.replace_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.replace_table, stretch=1)

        add_row = QHBoxLayout()
        self.new_from_input = QLineEdit()
        self.new_from_input.setPlaceholderText("置換する文字列（例: w）")
        self.new_to_input = QLineEdit()
        self.new_to_input.setPlaceholderText("読み（例: わら）")

        self.add_replace_btn = QPushButton(" 追加")
        self.add_replace_btn.setObjectName("PrimaryButton")
        self.add_replace_btn.setIcon(self._icon("add.png"))
        self.add_replace_btn.setIconSize(QSize(20, 20))

        add_row.addWidget(self.new_from_input)
        add_row.addWidget(self.new_to_input)
        add_row.addWidget(self.add_replace_btn)
        layout.addLayout(add_row)

        self._refresh_replace_table()

    # -------------------------
    # Stylesheet (Modern Discord Theme 2024+)
    # -------------------------
    def _apply_stylesheet(self):
        cb_on = (SCRIPT_DIR / "assets" / "checkbox_on.png").as_posix()
        cb_off = (SCRIPT_DIR / "assets" / "checkbox_off.png").as_posix()

        style = f"""
        QWidget {{
            font-family: "gg sans", "Segoe UI", "Hiragino Kaku Gothic ProN", Meiryo, sans-serif;
            font-size: 14px;
            color: #DBDEE1;
        }}
        QMainWindow, #CentralWidget {{
            background-color: #313338;
        }}

        QTabWidget::pane {{
            border: none;
            background-color: #313338;
            border-radius: 8px;
        }}
        QTabBar::tab {{
            background-color: #2B2D31;
            color: #949BA4;
            padding: 12px 24px;
            margin-right: 4px;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            font-weight: 800;
        }}
        QTabBar::tab:selected {{
            background-color: #313338;
            color: #F2F3F5;
            border-bottom: 3px solid #5865F2;
        }}
        QTabBar::tab:hover:!selected {{
            background-color: #393C43;
            color: #DBDEE1;
        }}

        QPushButton {{
            background-color: #4E5058;
            color: #FFFFFF;
            border: none;
            padding: 8px 20px;
            min-height: 24px;
            border-radius: 4px;
            font-weight: 800;
            font-size: 14px;
        }}
        QPushButton:hover {{
            background-color: #6D6F78;
        }}
        QPushButton:pressed {{
            background-color: #404249;
        }}
        QPushButton:disabled {{
            background-color: #1E1F22;
            color: #5C5E66;
        }}

        QPushButton#PrimaryButton, QPushButton#StartButton {{
            background-color: #5865F2;
        }}
        QPushButton#PrimaryButton:hover, QPushButton#StartButton:hover {{
            background-color: #4752C4;
        }}
        QPushButton#PrimaryButton:pressed, QPushButton#StartButton:pressed {{
            background-color: #3C45A5;
        }}

        QPushButton#StopButton, QPushButton[text="削除"] {{
            background-color: #DA373C;
        }}
        QPushButton#StopButton:hover, QPushButton[text="削除"]:hover {{
            background-color: #A12828;
        }}

        QLineEdit, QComboBox, QSpinBox {{
            background-color: #1E1F22;
            border: 1px solid #1E1F22;
            padding: 10px 14px;
            border-radius: 6px;
            color: #DBDEE1;
            min-height: 20px;
        }}
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
            border: 1px solid #5865F2;
        }}
        QComboBox::drop-down {{
            border: none;
            width: 30px;
        }}
        QComboBox::down-arrow {{
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #949BA4;
            margin-right: 10px;
        }}
        QComboBox QAbstractItemView {{
            background-color: #2B2D31;
            color: #DBDEE1;
            selection-background-color: #5865F2;
            border: 1px solid #1E1F22;
            border-radius: 4px;
        }}

        QTextEdit {{
            background-color: #1E1F22;
            border: 1px solid #1E1F22;
            border-radius: 8px;
            padding: 16px;
            color: #DBDEE1;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
        }}

        QTableWidget {{
            background-color: #2B2D31;
            border: none;
            border-radius: 8px;
            gridline-color: #1E1F22;
            color: #DBDEE1;
        }}
        QHeaderView::section {{
            background-color: #1E1F22;
            color: #949BA4;
            padding: 12px;
            border: none;
            font-weight: 900;
            font-size: 12px;
            text-transform: uppercase;
        }}
        QTableWidget::item {{
            padding: 8px;
            border-bottom: 1px solid #1E1F22;
        }}
        QTableWidget::item:selected {{
            background-color: rgba(88, 101, 242, 0.2);
            color: #FFFFFF;
        }}

        QScrollBar:vertical {{
            background: #2B2D31;
            width: 14px;
            margin: 0px;
            border-radius: 7px;
        }}
        QScrollBar::handle:vertical {{
            background: #1A1B1E;
            min-height: 30px;
            border-radius: 7px;
            margin: 2px;
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
        }}

        QSlider::groove:horizontal {{
            border-radius: 4px;
            height: 8px;
            background: #4E5058;
        }}
        QSlider::handle:horizontal {{
            background: #FFFFFF;
            width: 20px;
            height: 20px;
            margin: -6px 0;
            border-radius: 10px;
        }}
        QSlider::sub-page:horizontal {{
            background: #5865F2;
            border-radius: 4px;
        }}

        QCheckBox {{
            spacing: 8px;
        }}
        QCheckBox::indicator {{
            width: 26px;
            height: 26px;
            border: none;
            background-color: transparent;
        }}
        QCheckBox::indicator:unchecked {{
            image: url({cb_off});
        }}
        QCheckBox::indicator:checked {{
            image: url({cb_on});
        }}

        #FormContainer {{
            background-color: #2B2D31;
            border-radius: 8px;
        }}
        """
        self.setStyleSheet(style)

    # -------------------------
    # Signals / Callbacks
    # -------------------------
    def _connect_signals(self):
        self.start_button.clicked.connect(self._on_start_clicked)
        self.stop_button.clicked.connect(self._on_stop_clicked)
        self.speed_slider.valueChanged.connect(self._on_speed_changed)
        self.export_btn.clicked.connect(self._on_export)
        self.import_btn.clicked.connect(self._on_import)
        self.reload_members_btn.clicked.connect(self._refresh_map_table)
        self.add_user_btn.clicked.connect(self._on_add_user)
        self.add_pair_btn.clicked.connect(self._on_add_pair)
        self.save_token_btn.clicked.connect(self._on_save_token)
        self.add_replace_btn.clicked.connect(self._on_add_replace)

    def _on_start_clicked(self):
        self.set_status("CONNECTING")
        self.append_log("[INFO] Start pressed")

    def _on_stop_clicked(self):
        self.set_status("STOPPED")
        self.append_log("[INFO] Stop pressed")

    def _on_speed_changed(self, value: int):
        speed = value / 100.0
        self.speed_value_label.setText(f"{speed:.2f}x")
        try:
            self.config.set("speed", speed)
        except Exception:
            try:
                self.config.set("speed", speed, save=True)
            except Exception:
                pass

    def _on_export(self):
        default_path = _get_downloads_dir() / "VoxCord_config.json"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "設定をエクスポート",
            str(default_path),
            "JSON Files (*.json)",
        )
        if path:
            try:
                saved = self.config.export_to(path)
                QMessageBox.information(self, "完了", f"設定をエクスポートしました\n{saved}")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"エクスポートに失敗しました: {e}")

    def _on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "設定をインポート",
            str(SCRIPT_DIR),
            "JSON Files (*.json)",
        )
        if path:
            try:
                self.config.import_from(path)
                QMessageBox.information(self, "完了", "設定をインポートしました。UIをリロードします。")
                self._refresh_all()
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"インポートに失敗しました: {e}")

    def _on_add_user(self):
        user_id = self.add_user_id_input.text().strip()
        if not user_id:
            QMessageBox.warning(self, "入力エラー", "User ID を入力してください")
            return
        try:
            default_sp = int(self.config.get("default_speaker", 3))
            self.config.set_member_voice(user_id, default_sp)
            self.append_log(f"[INFO] ユーザーを追加しました: {user_id}")
            self._refresh_map_table()
            self.add_user_id_input.clear()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"追加に失敗しました: {e}")

    def _on_add_pair(self):
        guild = self.guild_input.text().strip()
        text = self.text_input.text().strip()
        voice = self.voice_input.text().strip()
        if not text or not voice:
            QMessageBox.warning(self, "入力エラー", "Text/Voice Channel ID を入力してください")
            return
        try:
            self.config.add_channel_pair(guild_id=guild, text_channel_id=text, voice_channel_id=voice)
            self.append_log(f"[INFO] チャンネルペアを追加しました: text={text} voice={voice}")
            self._refresh_channel_pairs()
            self.guild_input.clear()
            self.text_input.clear()
            self.voice_input.clear()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"追加に失敗しました: {e}")

    def _on_save_token(self):
        token = self.token_input.text().strip()
        if not token:
            QMessageBox.warning(self, "入力エラー", "Bot token を入力してください")
            return
        try:
            self.config.set_bot_token(token)
            QMessageBox.information(self, "完了", "Bot token を保存しました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存に失敗しました: {e}")

    def _on_add_replace(self):
        frm = self.new_from_input.text()
        to = self.new_to_input.text()
        if not frm:
            QMessageBox.warning(self, "入力エラー", "置換する文字列を入力してください")
            return
        try:
            try:
                self.config.add_replace_rule(frm, to)
            except Exception:
                rr = self.config.get("replace_rules", {})
                rr[frm] = to
                self.config.set("replace_rules", rr)
            self.append_log(f"[INFO] 置換ルール追加: {frm} -> {to}")
            self.new_from_input.clear()
            self.new_to_input.clear()
            self._refresh_replace_table()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"追加に失敗しました: {e}")

    # -------------------------
    # Refresh UI from config
    # -------------------------
    def _refresh_all(self):
        self.speed_slider.setValue(int(float(self.config.get("speed", 1.0)) * 100))
        self.token_input.setText(self.config.get_bot_token())
        self._refresh_channel_pairs()
        self._refresh_replace_table()
        self._refresh_map_table()

    def _refresh_channel_pairs(self):
        pairs = self.config.get_channel_pairs()
        self.pair_table.setRowCount(0)
        for p in pairs:
            row = self.pair_table.rowCount()
            self.pair_table.insertRow(row)

            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            chk_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            chk.setChecked(bool(p.get("enabled", True)))
            chk_layout.addWidget(chk)

            def make_toggle(t):
                return lambda state: self._toggle_pair_enabled(t, state)

            chk.stateChanged.connect(make_toggle(p.get("text_channel_id")))
            self.pair_table.setCellWidget(row, 0, chk_widget)

            self.pair_table.setItem(row, 1, QTableWidgetItem(str(p.get("guild_id", ""))))
            self.pair_table.setItem(row, 2, QTableWidgetItem(str(p.get("text_channel_id", ""))))
            self.pair_table.setItem(row, 3, QTableWidgetItem(str(p.get("voice_channel_id", ""))))

            rm_btn = QPushButton("削除")

            def make_rm(t):
                return lambda: self._remove_pair(t)

            rm_btn.clicked.connect(make_rm(p.get("text_channel_id")))
            self.pair_table.setCellWidget(row, 4, rm_btn)

    def _toggle_pair_enabled(self, text_channel_id, state):
        try:
            self.config.set_channel_pair_enabled(text_channel_id, bool(state))
            self.append_log(f"[INFO] チャンネル {text_channel_id} の有効状態を変更しました: {bool(state)}")
        except Exception:
            pass

    def _remove_pair(self, text_channel_id):
        ret = QMessageBox.question(self, "削除確認", f"Text ID {text_channel_id} を削除しますか？")
        if ret != QMessageBox.Yes:
            return
        try:
            self.config.remove_channel_pair(text_channel_id=text_channel_id)
            self._refresh_channel_pairs()
            self.append_log(f"[INFO] チャンネルペアを削除しました: {text_channel_id}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"削除に失敗しました: {e}")

    def _refresh_replace_table(self):
        rr = self.config.get_replace_rules()
        self.replace_table.setRowCount(0)
        for k, v in rr.items():
            row = self.replace_table.rowCount()
            self.replace_table.insertRow(row)
            self.replace_table.setItem(row, 0, QTableWidgetItem(str(k)))
            self.replace_table.setItem(row, 1, QTableWidgetItem(str(v)))
            rm_btn = QPushButton("削除")
            rm_btn.clicked.connect(lambda _, key=k: self._remove_replace(key))
            self.replace_table.setCellWidget(row, 2, rm_btn)

    def _remove_replace(self, key: str):
        confirm = QMessageBox.question(self, "削除確認", f"置換ルール '{key}' を削除しますか？")
        if confirm != QMessageBox.Yes:
            return
        try:
            self.config.remove_replace_rule(key)
            self._refresh_replace_table()
            self.append_log(f"[INFO] 置換ルールを削除しました: {key}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"削除に失敗しました: {e}")

    def _refresh_map_table(self):
        try:
            mvm = self.config.get("member_voice_map", {})
        except Exception:
            mvm = self.config.to_dict().get("member_voice_map", {})

        self.map_table.setRowCount(0)
        if not mvm:
            return

        for user_id, info in mvm.items():
            row = self.map_table.rowCount()
            self.map_table.insertRow(row)

            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            chk_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            enabled = bool(info.get("enabled", True))
            chk.setChecked(enabled)
            chk.stateChanged.connect(self._make_member_toggle(user_id))
            chk_layout.addWidget(chk)
            self.map_table.setCellWidget(row, 0, chk_widget)

            self.map_table.setItem(row, 1, QTableWidgetItem(str(user_id)))
            display_name = str(info.get("display_name", ""))

            self.map_table.setItem(row, 2, QTableWidgetItem(display_name))

            cb = QComboBox()
            self._populate_speakers_into(cb)
            speaker_id = int(info.get("speaker_id", self.config.get("default_speaker", 3)))
            idx = 0
            for i in range(cb.count()):
                if cb.itemData(i) == speaker_id:
                    idx = i
                    break
            cb.setCurrentIndex(idx)
            cb.currentIndexChanged.connect(self._make_member_speaker_changed(user_id, cb))
            self.map_table.setCellWidget(row, 3, cb)

            op_w = QWidget()
            op_layout = QHBoxLayout(op_w)
            op_layout.setContentsMargins(4, 4, 4, 4)
            op_layout.setSpacing(8)
            test_btn = QPushButton("▶ テスト")
            test_btn.setObjectName("SecondaryButton")
            rm_btn = QPushButton("削除")
            test_btn.clicked.connect(self._make_member_test(user_id))
            rm_btn.clicked.connect(self._make_member_remove(user_id))
            op_layout.addWidget(test_btn)
            op_layout.addWidget(rm_btn)
            self.map_table.setCellWidget(row, 4, op_w)

    def _make_member_toggle(self, user_id):
        def _toggle(state):
            try:
                mv = self.config.to_dict().get("member_voice_map", {})
                if str(user_id) in mv:
                    cur = mv[str(user_id)]
                    cur["enabled"] = bool(state)
                    self.config.set("member_voice_map", mv)
                    try:
                        self.config.set_member_voice(
                            str(user_id),
                            int(cur.get("speaker_id", self.config.get("default_speaker", 3))),
                            enabled=bool(state),
                        )
                    except Exception:
                        pass
                    self.append_log(f"[INFO] member {user_id} enabled={bool(state)}")
            except Exception:
                pass

        return _toggle

    def _make_member_speaker_changed(self, user_id, combobox: QComboBox):
        def _changed(idx):
            try:
                sp = combobox.itemData(idx)
                self.config.set_member_voice(str(user_id), int(sp))
                self.append_log(f"[INFO] member {user_id} speaker changed -> {sp}")
            except Exception:
                pass

        return _changed

    def _make_member_test(self, user_id):
        def _test():
            QMessageBox.information(self, "テスト", f"ユーザー {user_id} のテスト再生を要求します")
            self.append_log(f"[INFO] テスト再生要求: {user_id}")

        return _test

    def _make_member_remove(self, user_id):
        def _remove():
            confirm = QMessageBox.question(self, "削除確認", f"ユーザー {user_id} の設定を削除しますか？")
            if confirm != QMessageBox.Yes:
                return
            try:
                self.config.remove_member_voice(str(user_id))
                self._refresh_map_table()
                self.append_log(f"[INFO] member {user_id} を削除しました")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"削除に失敗しました: {e}")

        return _remove

    def append_log(self, text: str):
        self.log_view.append(f"<span style='color:#B5BAC1;'>{text}</span>")

    def set_status(self, status_text: str):
        self.status_label.setText(status_text)
        self._apply_status_style(status_text)

    def _apply_status_style(self, status_text: str):
        text = status_text.upper()
        if "STOP" in text:
            bg = "#DA373C"
        elif "CONNECT" in text or "START" in text:
            bg = "#FEE75C"
        elif "PLAY" in text or "OK" in text or "RUN" in text:
            bg = "#23A559"
        else:
            bg = "#80848E"

        color = "#000000" if bg == "#FEE75C" else "#FFFFFF"
        self.status_label.setStyleSheet(
            f"background-color: {bg}; color: {color}; border-radius: 6px; font-weight: 900; letter-spacing: 1px;"
        )

    def _on_config_changed(self, snapshot: Dict):
        try:
            self._refresh_all()
            self.append_log("[INFO] 設定が更新されました（外部）")
        except Exception:
            pass

    def show(self):
        self._refresh_all()
        super().show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    cfg = ConfigManager("config.json")
    gui = VoxCordGUI(cfg)
    gui.show()
    sys.exit(app.exec())