from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
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

        root_layout = QHBoxLayout()
        left_widget = QWidget(); left_widget.setLayout(left)
        right_widget = QWidget(); right_widget.setLayout(right)
        root_layout.addWidget(left_widget, 1)
        root_layout.addWidget(right_widget, 3)
        wrapper = QWidget(); wrapper.setLayout(root_layout)
        self.setCentralWidget(wrapper)

        self.refresh_events()

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
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите стартовый протокол",
            str(self.root / "data"),
            "Excel (*.xlsx *.xlsm)",
        )
        if not path:
            return
        self.service.import_startlist(Path(path))
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


def run_app() -> None:
    app = QApplication([])
    root = Path(__file__).resolve().parents[1]
    service = MeetService(root)
    window = MainWindow(service, root)
    window.show()
    app.exec()
    service.close()
