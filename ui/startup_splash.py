from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont, QPainter, QPixmap
from PyQt6.QtWidgets import QWidget

from version import VERSION


class PublishStartupSplash(QWidget):
    """Short, quiet Golden Fir brand splash."""

    def __init__(self, assets_dir: Path, parent=None):
        super().__init__(parent)
        self.assets_dir = Path(assets_dir)
        self._logo = QPixmap(str(self.assets_dir / "golden_fir_wordmark.png"))
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent, True)
        self.setWindowTitle("Master Report Center All In One")
        self.resize(620, 360)

    def show_centered(self, screen):
        geo = screen.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)
        self.show()
        self.raise_()
        self.activateWindow()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
        painter.fillRect(self.rect(), QColor("#F7F8F6"))

        if not self._logo.isNull():
            logo = self._logo.scaled(
                390, 145,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            painter.drawPixmap((self.width()-logo.width())//2, 48, logo)

        title_font = QFont("Segoe UI", 18)
        title_font.setBold(True)
        painter.setFont(title_font)
        painter.setPen(QColor("#173D2A"))
        painter.drawText(
            0, 205, self.width(), 30,
            int(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter),
            "Master Report Center All In One",
        )

        painter.setFont(QFont("Segoe UI", 9))
        painter.setPen(QColor("#66776E"))
        painter.drawText(
            0, 245, self.width(), 22,
            int(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter),
            f"Version {VERSION}",
        )
        painter.drawText(
            0, 283, self.width(), 20,
            int(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter),
            "Initializing workspace…",
        )
        painter.end()
