from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, Qt
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
        mark_absent_btn = QPushButton("Не явился")
        mark_absent_btn.clicked.connect(self.mark_absent)
        restore_btn = QPushButton("Вернуть в заплывы")
        restore_btn.clicked.connect(self.restore_swimmers)
        reseed_btn = QPushButton("Пересобрать")
        reseed_btn.clicked.connect(self.reseed_event)

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


def run_app() -> None:
    app = QApplication([])
    root = Path(__file__).resolve().parents[1]
    service = MeetService(root)
    window = MainWindow(service, root)
    window.show()
    app.exec()
    service.close()
