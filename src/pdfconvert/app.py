from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import QUrl

from .converter import ConversionError, SUPPORTED_EXTENSIONS, convert_and_merge


class FileList(QListWidget):
    files_dropped = Signal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QListWidget.ExtendedSelection)
        self.setAlternatingRowColors(True)
        self.setMinimumHeight(280)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            self.files_dropped.emit([url.toLocalFile() for url in event.mimeData().urls()])
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class Worker(QObject):
    progress = Signal(int, int, str)
    succeeded = Signal(str)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, files: list[Path], output: Path) -> None:
        super().__init__()
        self.files = files
        self.output = output

    def run(self) -> None:
        try:
            result = convert_and_merge(self.files, self.output, self.progress.emit)
            self.succeeded.emit(str(result))
        except (ConversionError, OSError, TimeoutError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            self.failed.emit(f"發生未預期錯誤：{exc}")
        finally:
            self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("文件合併轉 PDF")
        self.resize(760, 560)
        self.output_path: Path | None = None
        self.thread: QThread | None = None
        self.worker: Worker | None = None

        title = QLabel("拖曳 PPTX 或 DOCX 到下方")
        title.setStyleSheet("font-size: 22px; font-weight: 600;")
        hint = QLabel("清單順序就是 PDF 合併順序；可直接拖曳項目，或使用右側按鈕調整。")
        hint.setWordWrap(True)

        self.file_list = FileList()
        self.file_list.files_dropped.connect(self.add_files)
        self.file_list.model().rowsMoved.connect(self.refresh_numbers)

        add_button = QPushButton("加入檔案")
        add_button.clicked.connect(self.pick_files)
        remove_button = QPushButton("移除")
        remove_button.clicked.connect(self.remove_selected)
        up_button = QPushButton("上移")
        up_button.clicked.connect(lambda: self.move_selected(-1))
        down_button = QPushButton("下移")
        down_button.clicked.connect(lambda: self.move_selected(1))
        clear_button = QPushButton("清除全部")
        clear_button.clicked.connect(self.file_list.clear)

        side = QVBoxLayout()
        for button in (add_button, remove_button, up_button, down_button, clear_button):
            side.addWidget(button)
        side.addStretch()
        list_row = QHBoxLayout()
        list_row.addWidget(self.file_list, 1)
        list_row.addLayout(side)

        self.output_label = QLabel("尚未選擇輸出位置")
        output_button = QPushButton("選擇輸出 PDF")
        output_button.clicked.connect(self.pick_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_label, 1)
        output_row.addWidget(output_button)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.status = QLabel("")
        self.convert_button = QPushButton("轉檔並合併")
        self.convert_button.setMinimumHeight(44)
        self.convert_button.setStyleSheet("font-size: 16px; font-weight: 600;")
        self.convert_button.clicked.connect(self.start_conversion)

        layout = QVBoxLayout()
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addLayout(list_row, 1)
        layout.addLayout(output_row)
        layout.addWidget(self.status)
        layout.addWidget(self.progress)
        layout.addWidget(self.convert_button)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def add_files(self, names: list[str]) -> None:
        existing = {self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())}
        rejected: list[str] = []
        for name in names:
            path = Path(name).resolve()
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                if str(path) not in existing:
                    item = QListWidgetItem()
                    item.setData(Qt.UserRole, str(path))
                    item.setToolTip(str(path))
                    self.file_list.addItem(item)
                    existing.add(str(path))
            else:
                rejected.append(path.name)
        self.refresh_numbers()
        if rejected:
            QMessageBox.information(self, "略過不支援的檔案", "僅支援 PPTX 與 DOCX：\n" + "\n".join(rejected))

    def refresh_numbers(self, *_args) -> None:
        for index in range(self.file_list.count()):
            item = self.file_list.item(index)
            item.setText(f"{index + 1}.  {Path(item.data(Qt.UserRole)).name}")

    def pick_files(self) -> None:
        names, _ = QFileDialog.getOpenFileNames(self, "選擇文件", "", "Office 文件 (*.pptx *.docx)")
        self.add_files(names)

    def pick_output(self) -> None:
        name, _ = QFileDialog.getSaveFileName(self, "儲存合併 PDF", "合併文件.pdf", "PDF (*.pdf)")
        if name:
            self.output_path = Path(name).with_suffix(".pdf")
            self.output_label.setText(str(self.output_path))
            self.output_label.setToolTip(str(self.output_path))

    def remove_selected(self) -> None:
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self.refresh_numbers()

    def move_selected(self, offset: int) -> None:
        row = self.file_list.currentRow()
        target = row + offset
        if row < 0 or target < 0 or target >= self.file_list.count():
            return
        item = self.file_list.takeItem(row)
        self.file_list.insertItem(target, item)
        self.file_list.setCurrentRow(target)
        self.refresh_numbers()

    def start_conversion(self) -> None:
        if not self.file_list.count():
            QMessageBox.warning(self, "沒有檔案", "請先加入 PPTX 或 DOCX 檔案。")
            return
        if self.output_path is None:
            self.pick_output()
            if self.output_path is None:
                return
        files = [Path(self.file_list.item(i).data(Qt.UserRole)) for i in range(self.file_list.count())]
        self.convert_button.setEnabled(False)
        self.file_list.setEnabled(False)
        self.progress.setRange(0, len(files))
        self.progress.setValue(0)
        self.progress.setVisible(True)

        self.thread = QThread(self)
        self.worker = Worker(files, self.output_path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.succeeded.connect(self.conversion_done)
        self.worker.failed.connect(lambda message: QMessageBox.critical(self, "轉檔失敗", message))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.reset_ui)
        self.thread.start()

    def update_progress(self, value: int, maximum: int, message: str) -> None:
        self.progress.setRange(0, maximum)
        self.progress.setValue(value)
        self.status.setText(message)

    def conversion_done(self, output: str) -> None:
        self.progress.setValue(self.progress.maximum())
        self.status.setText("完成")
        box = QMessageBox(self)
        box.setWindowTitle("轉檔完成")
        box.setText(f"PDF 已儲存至：\n{output}")
        open_button = box.addButton("開啟 PDF", QMessageBox.AcceptRole)
        box.addButton("關閉", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is open_button:
            QDesktopServices.openUrl(QUrl.fromLocalFile(output))

    def reset_ui(self) -> None:
        self.convert_button.setEnabled(True)
        self.file_list.setEnabled(True)
        self.thread = None
        self.worker = None


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("文件合併轉 PDF")
    window = MainWindow()
    window.show()
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()

