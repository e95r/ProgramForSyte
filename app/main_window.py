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
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QDialog,
    QDialogButtonBox,
    QComboBox,
    QPlainTextEdit,
    QPushButton,
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

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["ID", "ФИО", "Год", "Команда", "Время", "Заплыв", "Статус"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        import_btn = QPushButton("Импорт Excel")
        import_btn.clicked.connect(self.import_excel)
        backup_btn = QPushButton("Бэкап БД")
        backup_btn.clicked.connect(self.make_backup)
        dns_btn = QPushButton("Пометить DNS и пересобрать")
        dns_btn.clicked.connect(self.mark_dns)
        input_results_btn = QPushButton("Ввод результатов заплыва")
        input_results_btn.clicked.connect(self.open_results_input)
        start_protocol_btn = QPushButton("Стартовый протокол")
        start_protocol_btn.clicked.connect(self.open_start_protocol)
        show_protocol_btn = QPushButton("Протокол дистанции")
        show_protocol_btn.clicked.connect(self.open_event_protocol)
        final_btn = QPushButton("Подвести итоги соревнования")
        final_btn.clicked.connect(self.open_final_protocol)
        self.protocol_sort = QComboBox()
        self.protocol_sort.addItem("По результатам", "result")
        self.protocol_sort.addItem("По заплывам", "heat")

        left = QVBoxLayout()
        left.addWidget(QLabel("Дистанции"))
        left.addWidget(self.events_list)
        left.addWidget(import_btn)
        left.addWidget(backup_btn)

        right = QVBoxLayout()
        right.addWidget(self.search_input)
        right.addWidget(self.full_reseed)
        right.addWidget(self.table)
        right.addWidget(dns_btn)
        right.addWidget(input_results_btn)
        right.addWidget(start_protocol_btn)
        right.addWidget(QLabel("Сортировка/группировка протокола"))
        right.addWidget(self.protocol_sort)
        right.addWidget(show_protocol_btn)
        right.addWidget(final_btn)

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
                s.status,
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

    def mark_dns(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            return
        selected = self.table.selectionModel().selectedRows()
        swimmer_ids = [int(self.table.item(idx.row(), 0).text()) for idx in selected]
        if not swimmer_ids:
            QMessageBox.warning(self, "DNS", "Выберите хотя бы одного спортсмена")
            return

        mode = "full" if self.full_reseed.isChecked() else "soft"
        if mode == "full":
            reply = QMessageBox.question(
                self,
                "Полный пересев",
                "Полный пересев изменит номера заплывов. Продолжить?",
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.service.mark_dns(event_id, swimmer_ids, mode=mode)
        self.load_swimmers()

    def make_backup(self) -> None:
        backup = self.service.create_backup(reason="manual")
        if backup:
            QMessageBox.information(self, "Бэкап", f"Сохранено: {backup}")
        else:
            QMessageBox.warning(self, "Бэкап", "База ещё не создана")

    def open_results_input(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            QMessageBox.warning(self, "Результаты", "Сначала выберите дистанцию")
            return
        dialog = ResultsInputDialog(self.service, event_id, self)
        dialog.exec()

    def open_event_protocol(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            QMessageBox.warning(self, "Протокол", "Сначала выберите дистанцию")
            return
        title = "Протокол дистанции"
        sort_mode = self.protocol_sort.currentData()
        text = self.service.build_event_protocol_text(event_id, sort_mode=sort_mode)

        def saver(path: Path) -> None:
            path.write_text(text, encoding="utf-8")
            self.service.repo.log("save_event_protocol_manual", str(path))

        default_name = f"event-{event_id}-protocol.txt"
        dialog = ProtocolDialog(title, text, saver, default_name, self)
        dialog.exec()

    def open_start_protocol(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            QMessageBox.warning(self, "Стартовый протокол", "Сначала выберите дистанцию")
            return
        title = "Стартовый протокол дистанции"
        text = self.service.build_start_protocol_text(event_id)

        def saver(path: Path) -> None:
            path.write_text(text, encoding="utf-8")
            self.service.repo.log("save_start_protocol_manual", str(path))

        default_name = f"event-{event_id}-start-protocol.txt"
        dialog = ProtocolDialog(title, text, saver, default_name, self)
        dialog.exec()

    def open_final_protocol(self) -> None:
        sort_mode = self.protocol_sort.currentData()
        default_path = self.service.save_final_protocol(sort_mode=sort_mode)
        title = "Итоговый протокол соревнования"
        text = self.service.build_final_protocol_text(sort_mode=sort_mode)

        def saver(path: Path) -> None:
            path.write_text(text, encoding="utf-8")
            self.service.repo.log("save_final_protocol_manual", str(path))

        dialog = ProtocolDialog(title, text, saver, "final-protocol.txt", self)
        QMessageBox.information(self, "Итоги", f"Итоговый протокол сохранён: {default_path}")
        dialog.exec()


class ResultsInputDialog(QDialog):
    def __init__(self, service: MeetService, event_id: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.service = service
        self.event_id = event_id
        self.setWindowTitle("Ввод результатов заплыва")
        self.resize(900, 600)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels([
            "ID",
            "ФИО",
            "Заплыв",
            "Дорожка",
            "Статус заявки",
            "Заявка",
            "Результат",
            "Статус результата",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Close)
        buttons.accepted.connect(self.save)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(buttons)
        self.setLayout(layout)
        self.load_data()

    def load_data(self) -> None:
        swimmers = self.service.repo.list_swimmers(self.event_id)
        self.table.setRowCount(len(swimmers))
        for row_idx, swimmer in enumerate(swimmers):
            static_values = [
                str(swimmer.id),
                swimmer.full_name,
                str(swimmer.heat or "-"),
                str(swimmer.lane or "-"),
                swimmer.status,
                swimmer.seed_time_raw or "",
            ]
            for col_idx, value in enumerate(static_values):
                item = QTableWidgetItem(value)
                if col_idx < 6:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_idx, col_idx, item)
            self.table.setItem(row_idx, 6, QTableWidgetItem(swimmer.result_time_raw or ""))
            status_combo = QComboBox()
            status_combo.addItems(["OK", "DQ", "DNS", "DNF"])
            status_combo.setCurrentText(swimmer.result_status or "OK")
            self.table.setCellWidget(row_idx, 7, status_combo)

    def save(self) -> None:
        payload: list[dict] = []
        for row in range(self.table.rowCount()):
            swimmer_id = int(self.table.item(row, 0).text())
            result_text = self.table.item(row, 6).text().strip()
            status_combo = self.table.cellWidget(row, 7)
            result_status = status_combo.currentText() if isinstance(status_combo, QComboBox) else "OK"
            payload.append(
                {
                    "swimmer_id": swimmer_id,
                    "result_time_raw": result_text,
                    "result_status": result_status,
                }
            )
        self.service.save_event_results(self.event_id, payload)
        QMessageBox.information(self, "Результаты", "Результаты сохранены")
        self.accept()


class ProtocolDialog(QDialog):
    def __init__(self, title: str, text: str, saver, default_name: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(900, 650)
        self.text = text
        self.saver = saver
        self.default_name = default_name

        self.viewer = QPlainTextEdit()
        self.viewer.setReadOnly(True)
        self.viewer.setPlainText(text)

        print_btn = QPushButton("Печать")
        print_btn.clicked.connect(self.print_protocol)
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_protocol)
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)

        buttons = QHBoxLayout()
        buttons.addWidget(print_btn)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)

        layout = QVBoxLayout()
        layout.addWidget(self.viewer)
        layout.addLayout(buttons)
        self.setLayout(layout)

    def print_protocol(self) -> None:
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            document = QTextDocument(self.text)
            document.print(printer)

    def save_protocol(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить протокол",
            str(Path.home() / self.default_name),
            "Text Files (*.txt);;All Files (*)",
        )
        if not path:
            return
        target = Path(path)
        self.saver(target)
        QMessageBox.information(self, "Сохранение", f"Сохранено: {target}")


def run_app() -> None:
    app = QApplication([])
    root = Path(__file__).resolve().parents[1]
    service = MeetService(root)
    window = MainWindow(service, root)
    window.show()
    app.exec()
    service.close()
