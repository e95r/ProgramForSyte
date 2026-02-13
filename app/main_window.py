from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, Qt
from PySide6.QtGui import QTextDocument
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.service import MeetService
from core.excel_importer import ExcelImportError


class MainWindow(QMainWindow):
    def __init__(self, service: MeetService, root: Path):
        super().__init__()
        self.service = service
        self.root = root
        self.setWindowTitle("Swim Meet MVP A+B")
        self.resize(1100, 700)

        self.events_list = QListWidget()
        self.events_list.currentRowChanged.connect(self.load_swimmers)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по ФИО")
        self.search_input.textChanged.connect(lambda _v: self.load_swimmers())

        self.full_reseed = QCheckBox("Полный пересев")
        self.full_reseed.setToolTip("Если включено — полностью пересчитать заплывы по заявочному времени")

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(["ID", "ФИО", "Год", "Команда", "Время", "Заплыв", "Статус", "Результат"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        import_btn = QPushButton("Импорт Excel")
        import_btn.clicked.connect(self.import_excel)
        backup_btn = QPushButton("Бэкап БД")
        backup_btn.clicked.connect(self.make_backup)
        mark_absent_btn = QPushButton("Не явился")
        mark_absent_btn.clicked.connect(self.mark_absent)
        restore_btn = QPushButton("Вернуть в заплывы")
        restore_btn.clicked.connect(self.restore_swimmers)
        reseed_btn = QPushButton("Пересобрать")
        reseed_btn.clicked.connect(self.reseed_event)
        result_entry_btn = QPushButton("Ввод результатов")
        result_entry_btn.clicked.connect(self.open_results_entry)
        event_protocol_btn = QPushButton("Протокол дистанции")
        event_protocol_btn.clicked.connect(self.open_event_protocol)
        final_protocol_btn = QPushButton("Подвести итоги")
        final_protocol_btn.clicked.connect(self.open_final_protocol)

        left = QVBoxLayout()
        left.addWidget(QLabel("Дистанции"))
        left.addWidget(self.events_list)
        left.addWidget(import_btn)
        left.addWidget(backup_btn)

        right = QVBoxLayout()
        right.addWidget(self.search_input)
        right.addWidget(self.full_reseed)
        right.addWidget(self.table)
        actions = QHBoxLayout()
        actions.addWidget(mark_absent_btn)
        actions.addWidget(restore_btn)
        actions.addWidget(reseed_btn)
        actions.addWidget(result_entry_btn)
        actions.addWidget(event_protocol_btn)
        actions.addWidget(final_protocol_btn)
        right.addLayout(actions)

        root_layout = QHBoxLayout()
        left_widget = QWidget(); left_widget.setLayout(left)
        right_widget = QWidget(); right_widget.setLayout(right)
        root_layout.addWidget(left_widget, 1)
        root_layout.addWidget(right_widget, 3)
        wrapper = QWidget(); wrapper.setLayout(root_layout)
        self.setCentralWidget(wrapper)

        self.refresh_events()

    def _file_debug_message(self, file_path: Path) -> str:
        exists = file_path.exists()
        size = file_path.stat().st_size if exists else 0
        suffix = file_path.suffix.lower()
        return (
            f"Selected: {file_path}\n"
            f"Exists: {exists}\n"
            f"Size: {size}\n"
            f"Suffix: {suffix}"
        )

    def refresh_events(self) -> None:
        self.events_list.clear()
        for event in self.service.repo.list_events():
            self.events_list.addItem(f"{event.id}. {event.name}")
        if self.events_list.count() > 0:
            self.events_list.setCurrentRow(0)

    def current_event_id(self) -> int | None:
        item = self.events_list.currentItem()
        if not item:
            return None
        return int(item.text().split(".", 1)[0])

    def load_swimmers(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            self.table.setRowCount(0)
            return
        swimmers = self.service.repo.list_swimmers(event_id, self.search_input.text())
        self.table.setRowCount(len(swimmers))
        for row_idx, s in enumerate(swimmers):
            values = [
                str(s.id),
                s.full_name,
                str(s.birth_year or ""),
                s.team or "",
                s.seed_time_raw or "",
                f"{s.heat or '-'} / {s.lane or '-'}",
                self._status_label(s.status),
                s.result_time_raw or "",
            ]
            for col_idx, val in enumerate(values):
                cell = QTableWidgetItem(val)
                if s.status == "DNS":
                    cell.setForeground(Qt.GlobalColor.darkGray)
                self.table.setItem(row_idx, col_idx, cell)

    def import_excel(self) -> None:
        settings = QSettings("ProgramForSyte", "SwimMeet")
        last_file = settings.value("last_opened_file", "", type=str)
        default_directory = str(Path(last_file).parent) if last_file else str(Path.home())

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите стартовый протокол",
            default_directory,
        )
        if not path:
            return

        selected_path = Path(path)
        settings.setValue("last_opened_file", str(selected_path))

        debug_message = self._file_debug_message(selected_path)
        print(debug_message)
        QMessageBox.information(self, "Диагностика выбора файла", debug_message)

        try:
            self.service.import_startlist(selected_path)
        except ExcelImportError as exc:
            QMessageBox.warning(self, "Ошибка импорта", str(exc))
            return

        self.refresh_events()
        QMessageBox.information(self, "Импорт", "Файл импортирован")

    def selected_swimmer_ids(self) -> list[int]:
        selected = self.table.selectionModel().selectedRows()
        return [int(self.table.item(idx.row(), 0).text()) for idx in selected]

    def _status_label(self, status: str) -> str:
        if status == "DNS":
            return "Не явился"
        return "В заявке"

    def reseed_mode(self) -> str:
        mode = "full" if self.full_reseed.isChecked() else "soft"
        if mode == "full":
            reply = QMessageBox.question(
                self,
                "Полный пересев",
                "Полный пересев изменит номера заплывов. Продолжить?",
            )
            if reply != QMessageBox.StandardButton.Yes:
                return ""
        return mode

    def mark_absent(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return
        swimmer_ids = self.selected_swimmer_ids()
        if not swimmer_ids:
            QMessageBox.warning(self, "Не явился", "Выберите хотя бы одного спортсмена")
            return

        self.service.mark_dns(event_id, swimmer_ids)
        self.load_swimmers()

    def restore_swimmers(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return
        swimmer_ids = self.selected_swimmer_ids()
        if not swimmer_ids:
            QMessageBox.warning(self, "Вернуть в заплывы", "Выберите хотя бы одного спортсмена")
            return

        mode = self.reseed_mode()
        if not mode:
            return

        self.service.restore_swimmers(event_id, swimmer_ids, mode=mode)
        self.load_swimmers()

    def reseed_event(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return

        mode = self.reseed_mode()
        if not mode:
            return

        self.service.reseed_event(event_id, mode=mode)
        self.load_swimmers()

    def make_backup(self) -> None:
        backup = self.service.create_backup(reason="manual")
        if backup:
            QMessageBox.information(self, "Бэкап", f"Сохранено: {backup}")
        else:
            QMessageBox.warning(self, "Бэкап", "База ещё не создана")

    def open_results_entry(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return
        dialog = ResultsEntryDialog(self.service, event_id, self)
        if dialog.exec():
            self.load_swimmers()

    def open_event_protocol(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return
        dialog = ProtocolDialog(
            self.service,
            title="Протокол дистанции",
            build_html=lambda grouped: self.service.build_event_protocol(event_id, grouped=grouped),
            self_parent=self,
        )
        dialog.exec()

    def open_final_protocol(self) -> None:
        dialog = ProtocolDialog(
            self.service,
            title="Итоговый протокол соревнований",
            build_html=lambda grouped: self.service.build_final_protocol(grouped=grouped),
            self_parent=self,
        )
        dialog.exec()


class ResultsEntryDialog(QDialog):
    def __init__(self, service: MeetService, event_id: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.service = service
        self.event_id = event_id
        self.setWindowTitle("Ввод результатов заплыва")
        self.resize(900, 600)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["ID", "Заплыв", "ФИО", "Команда", "Заявка", "Результат", "Отметка"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

        save_btn = QPushButton("Сохранить результаты")
        save_btn.clicked.connect(self.save_results)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(save_btn)
        self.setLayout(layout)
        self.load_rows()

    def load_rows(self) -> None:
        swimmers = self.service.repo.list_swimmers(self.event_id)
        self.table.setRowCount(len(swimmers))
        for row_idx, s in enumerate(swimmers):
            values = [
                str(s.id),
                f"{s.heat or '-'} / {s.lane or '-'}",
                s.full_name,
                s.team or "",
                s.seed_time_raw or "",
            ]
            for col_idx, val in enumerate(values):
                item = QTableWidgetItem(val)
                if col_idx < 5:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_idx, col_idx, item)

            self.table.setItem(row_idx, 5, QTableWidgetItem(s.result_time_raw or ""))
            self.table.setItem(row_idx, 6, QTableWidgetItem(s.result_mark or ""))

    def save_results(self) -> None:
        payload: list[dict[str, str]] = []
        for row in range(self.table.rowCount()):
            payload.append(
                {
                    "swimmer_id": self.table.item(row, 0).text(),
                    "result_time_raw": self.table.item(row, 5).text() if self.table.item(row, 5) else "",
                    "result_mark": self.table.item(row, 6).text() if self.table.item(row, 6) else "",
                }
            )
        self.service.save_event_results(self.event_id, payload)
        QMessageBox.information(self, "Результаты", "Результаты сохранены")
        self.accept()


class ProtocolDialog(QDialog):
    def __init__(self, service: MeetService, title: str, build_html, self_parent: QWidget | None = None):
        super().__init__(self_parent)
        self.service = service
        self.build_html = build_html
        self.setWindowTitle(title)
        self.resize(1000, 700)

        self.group_by_heat = QCheckBox("Группировать по заплывам/дорожкам")
        self.group_by_heat.setChecked(True)
        self.group_by_heat.stateChanged.connect(self.refresh_html)

        self.viewer = QTextEdit()
        self.viewer.setReadOnly(True)

        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_html)
        print_btn = QPushButton("Печать")
        print_btn.clicked.connect(self.print_protocol)
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_protocol)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self.group_by_heat)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(print_btn)
        toolbar.addWidget(save_btn)

        layout = QVBoxLayout()
        layout.addLayout(toolbar)
        layout.addWidget(self.viewer)
        self.setLayout(layout)
        self.refresh_html()

    def current_html(self) -> str:
        return self.build_html(self.group_by_heat.isChecked())

    def refresh_html(self) -> None:
        self.viewer.setHtml(self.current_html())

    def print_protocol(self) -> None:
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            doc = QTextDocument()
            doc.setHtml(self.current_html())
            doc.print(printer)

    def save_protocol(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить протокол",
            str(Path.home() / "protocol.html"),
            "HTML (*.html);;Text (*.txt)",
        )
        if not path:
            return
        text = self.current_html()
        Path(path).write_text(text, encoding="utf-8")
        QMessageBox.information(self, "Сохранение", f"Протокол сохранён: {path}")


def run_app() -> None:
    app = QApplication([])
    root = Path(__file__).resolve().parents[1]
    service = MeetService(root)
    window = MainWindow(service, root)
    window.show()
    app.exec()
    service.close()
