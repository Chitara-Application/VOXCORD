from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QLabel,
    QProgressBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class LoadingWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("VoxCord - Loading")
        self.setWindowFlags(
            Qt.Window
            | Qt.FramelessWindowHint
            | Qt.WindowStaysOnTopHint
        )
        self.setMinimumSize(560, 360)

        self._build_ui()
        self._apply_style()
        self.center_on_screen()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        self.card = QFrame()
        self.card.setObjectName("Card")
        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(22, 22, 22, 22)
        card_layout.setSpacing(12)

        self.title_label = QLabel("VoxCord")
        title_font = QFont()
        title_font.setPointSize(22)
        title_font.setBold(True)
        self.title_label.setFont(title_font)

        self.status_label = QLabel("起動準備中…")
        status_font = QFont()
        status_font.setPointSize(14)
        status_font.setBold(True)
        self.status_label.setFont(status_font)
        self.status_label.setWordWrap(True)

        self.detail_label = QLabel("")
        self.detail_label.setWordWrap(True)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setTextVisible(False)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("起動ログが表示されます")
        self.log_view.setMinimumHeight(140)

        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.status_label)
        card_layout.addWidget(self.detail_label)
        card_layout.addWidget(self.progress)
        card_layout.addWidget(self.log_view)

        root.addWidget(self.card)

    def _apply_style(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #0f1115;
                color: #e8e8e8;
            }
            QFrame#Card {
                background: #171a21;
                border: 1px solid #2a2f3a;
                border-radius: 16px;
            }
            QLabel {
                color: #e8e8e8;
            }
            QProgressBar {
                border: 1px solid #2a2f3a;
                border-radius: 6px;
                background: #10131a;
                height: 12px;
            }
            QProgressBar::chunk {
                border-radius: 6px;
                background: #4a90e2;
            }
            QTextEdit {
                background: #10131a;
                border: 1px solid #2a2f3a;
                border-radius: 10px;
                color: #dfe6f3;
                font-family: Consolas, "Yu Gothic UI", monospace;
                font-size: 10pt;
            }
            """
        )

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)

    def set_status(self, text: str):
        self.status_label.setText(text)

    def set_detail(self, text: str):
        self.detail_label.setText(text)

    def set_progress_indeterminate(self):
        self.progress.setRange(0, 0)

    def set_progress_value(self, value: int):
        self.progress.setRange(0, 100)
        self.progress.setValue(max(0, min(100, value)))

    def append_log(self, text: str):
        self.log_view.append(text)
        self.log_view.verticalScrollBar().setValue(
            self.log_view.verticalScrollBar().maximum()
        )
