from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, QSizeF, Qt
from PySide6.QtGui import QFont, QPageLayout, QPageSize, QTextDocument
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QSpinBox,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.excel_importer import ExcelImportError
from core.models import Secretary
from core.service import MeetService


UI_SETTINGS_ORG = "ProgramForSyte"
UI_SETTINGS_APP = "SwimMeet"
DEFAULT_FONT_FAMILY = "Arial"
DEFAULT_FONT_SIZE = 10
DEFAULT_THEME = "light"
FONT_FAMILIES = ["Arial", "Times New Roman", "Verdana", "Tahoma", "Calibri", "Segoe UI"]
THEME_LABELS = {
    "light": "Светлая",
    "dark": "Тёмная",
}


def load_ui_preferences() -> dict[str, str | int]:
    settings = QSettings(UI_SETTINGS_ORG, UI_SETTINGS_APP)
    theme = settings.value("ui/theme", DEFAULT_THEME, type=str)
    if theme not in THEME_LABELS:
        theme = DEFAULT_THEME
    font_family = settings.value("ui/font_family", DEFAULT_FONT_FAMILY, type=str)
    if font_family not in FONT_FAMILIES:
        font_family = DEFAULT_FONT_FAMILY
    font_size = settings.value("ui/font_size", DEFAULT_FONT_SIZE, type=int)
    if font_size < 8 or font_size > 24:
        font_size = DEFAULT_FONT_SIZE
    return {
        "theme": theme,
        "font_family": font_family,
        "font_size": font_size,
    }


def build_theme_stylesheet(theme: str) -> str:
    if theme == "dark":
        return """
        QWidget {
            background-color: #14161a;
            color: #eef2f7;
            font-size: 13px;
        }
        QMainWindow, QDialog {
            background-color: #14161a;
        }
        QLabel#sectionTitle {
            color: #8ab4ff;
            font-size: 14px;
            font-weight: 700;
            letter-spacing: 0.3px;
        }
        QWidget#sidebarCard, QWidget#contentCard {
            background-color: #1c1f24;
            border: 1px solid #2a2f38;
            border-radius: 12px;
        }
        QPushButton {
            background-color: #262b35;
            border: 1px solid #303746;
            padding: 8px 12px;
            border-radius: 8px;
            font-weight: 600;
        }
        QPushButton:hover {
            background-color: #2e3441;
        }
        QPushButton:pressed {
            background-color: #20242d;
        }
        QPushButton#primaryButton {
            background-color: #2f80ed;
            border: 1px solid #2f80ed;
            color: #ffffff;
        }
        QPushButton#primaryButton:hover {
            background-color: #3a8cf6;
        }
        QLineEdit, QTextEdit, QListWidget, QTableWidget, QComboBox, QSpinBox {
            background-color: #1b1f27;
            color: #eef2f7;
            border: 1px solid #323a4b;
            border-radius: 8px;
            padding: 6px;
            selection-background-color: #2f80ed;
            selection-color: #ffffff;
        }
        QLineEdit:focus, QTextEdit:focus, QListWidget:focus, QTableWidget:focus, QComboBox:focus, QSpinBox:focus {
            border: 1px solid #5ea1f5;
        }
        QHeaderView::section {
            background-color: #232833;
            color: #eef2f7;
            border: none;
            border-bottom: 1px solid #303746;
            padding: 8px 6px;
            font-weight: 600;
        }
        QTableWidget::item:selected {
            background-color: #2f80ed;
            color: #ffffff;
        }
        QTableWidget {
            gridline-color: #2a2f38;
        }
        QTabWidget::pane {
            border: 1px solid #2a2f38;
            border-radius: 10px;
            padding: 6px;
        }
        QTabBar::tab {
            background-color: #222834;
            color: #eef2f7;
            padding: 8px 12px;
            border: 1px solid #303746;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
        }
        QTabBar::tab:selected {
            background-color: #2f80ed;
            border-color: #2f80ed;
        }
        QSplitter::handle {
            background-color: #222834;
            width: 10px;
        }
        """
    return """
    QWidget {
        background-color: #f2f5fa;
        color: #1b2330;
        font-size: 13px;
    }
    QMainWindow, QDialog {
        background-color: #f2f5fa;
    }
    QLabel#sectionTitle {
        color: #1a5fd0;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    QWidget#sidebarCard, QWidget#contentCard {
        background-color: #ffffff;
        border: 1px solid #d7deeb;
        border-radius: 12px;
    }
    QPushButton {
        background-color: #ffffff;
        border: 1px solid #c8d3e6;
        padding: 8px 12px;
        border-radius: 8px;
        font-weight: 600;
    }
    QPushButton:hover {
        background-color: #eef4ff;
    }
    QPushButton:pressed {
        background-color: #dde9ff;
    }
    QPushButton#primaryButton {
        background-color: #2f80ed;
        border: 1px solid #2f80ed;
        color: #ffffff;
    }
    QPushButton#primaryButton:hover {
        background-color: #428ff5;
    }
    QLineEdit, QTextEdit, QListWidget, QTableWidget, QComboBox, QSpinBox {
        background-color: #ffffff;
        color: #1b2330;
        border: 1px solid #cad6eb;
        border-radius: 8px;
        padding: 6px;
        selection-background-color: #5b9dff;
        selection-color: #ffffff;
    }
    QLineEdit:focus, QTextEdit:focus, QListWidget:focus, QTableWidget:focus, QComboBox:focus, QSpinBox:focus {
        border: 1px solid #5b9dff;
    }
    QHeaderView::section {
        background-color: #f7f9fe;
        color: #1b2330;
        border: none;
        border-bottom: 1px solid #d9e1ee;
        padding: 8px 6px;
        font-weight: 600;
    }
    QTableWidget::item:selected {
        background-color: #5b9dff;
        color: #ffffff;
    }
    QTableWidget::item:selected:active {
        background-color: #2f80ed;
        color: #ffffff;
    }
    QTableWidget {
        gridline-color: #e1e7f3;
    }
    QTabWidget::pane {
        border: 1px solid #d4deee;
        border-radius: 10px;
        padding: 6px;
    }
    QTabBar::tab {
        background-color: #eaf1ff;
        color: #1b2330;
        padding: 8px 12px;
        border: 1px solid #cdd9ef;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
    }
    QTabBar::tab:selected {
        background-color: #2f80ed;
        border-color: #2f80ed;
        color: #ffffff;
    }
    QSplitter::handle {
        background-color: #dbe4f2;
        width: 10px;
    }
    """


def apply_ui_preferences(app: QApplication) -> dict[str, str | int]:
    preferences = load_ui_preferences()
    app.setFont(QFont(str(preferences["font_family"]), int(preferences["font_size"])))
    app.setStyleSheet(build_theme_stylesheet(str(preferences["theme"])))
    return preferences


class MainWindow(QMainWindow):
    @staticmethod
    def _apply_heat_spans(table: QTableWidget, rows: list[tuple[int | None, int]]) -> None:
        if table.rowCount() == 0:
            return
        table.clearSpans()
        start_row = 0
        while start_row < len(rows):
            heat, _lane = rows[start_row]
            end_row = start_row + 1
            while end_row < len(rows) and rows[end_row][0] == heat:
                end_row += 1
            span = end_row - start_row
            heat_item = table.item(start_row, 0)
            if heat_item is None:
                heat_item = QTableWidgetItem(str(heat or "-"))
                table.setItem(start_row, 0, heat_item)
            heat_item.setText(str(heat or "-"))
            heat_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if span > 1:
                table.setSpan(start_row, 0, span, 1)
            start_row = end_row

    def __init__(self, service: MeetService, root: Path, current_secretary: Secretary):
        super().__init__()
        self.service = service
        self.root = root
        self.current_secretary = current_secretary
        self.setWindowTitle(f"Swim Meet MVP A+B — секретарь: {current_secretary.display_name}")
        self.resize(1100, 820)

        self.events_list = QListWidget()
        self.events_list.currentRowChanged.connect(self.load_swimmers)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по ФИО")
        self.search_input.textChanged.connect(lambda _v: self.load_swimmers())

        self.full_reseed = QCheckBox("Полный пересев")
        self.full_reseed.setToolTip("Если включено — полностью пересчитать заплывы по заявочному времени")

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            ["Дистанция", "Заплыв", "Дорожка", "ФИО", "Год", "Команда", "Время", "Статус", "Результат"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)

        import_btn = QPushButton("Импорт Excel")
        import_btn.setObjectName("primaryButton")
        import_btn.clicked.connect(self.import_excel)
        backup_btn = QPushButton("Бэкап БД")
        backup_btn.clicked.connect(self.make_backup)
        settings_btn = QPushButton("Настройки")
        settings_btn.clicked.connect(self.open_settings)
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
        final_protocol_btn.setObjectName("primaryButton")
        final_protocol_btn.clicked.connect(self.open_final_protocol)

        secretary_label = QLabel(
            f"Вошёл секретарь: <b>{current_secretary.display_name}</b> ({current_secretary.username})<br>"
            f"Всего зарегистрировано секретарей: {self.service.secretary_count()}"
        )
        secretary_label.setWordWrap(True)

        left = QVBoxLayout()
        left.setContentsMargins(16, 16, 16, 16)
        left.setSpacing(12)
        events_label = QLabel("Дистанции")
        events_label.setObjectName("sectionTitle")
        left.addWidget(events_label)
        left.addWidget(self.events_list, 1)
        left.addWidget(import_btn)
        left.addWidget(backup_btn)
        left.addWidget(settings_btn)

        right = QVBoxLayout()
        right.setContentsMargins(16, 16, 16, 16)
        right.setSpacing(12)
        right.addWidget(secretary_label)
        right.addWidget(self.search_input)
        right.addWidget(self.full_reseed)
        right.addWidget(self.table)
        actions_top = QHBoxLayout()
        actions_top.setSpacing(8)
        actions_top.addWidget(mark_absent_btn)
        actions_top.addWidget(restore_btn)
        actions_top.addWidget(reseed_btn)
        actions_bottom = QHBoxLayout()
        actions_bottom.setSpacing(8)
        actions_bottom.addWidget(result_entry_btn)
        actions_bottom.addWidget(event_protocol_btn)
        actions_bottom.addWidget(final_protocol_btn)
        right.addLayout(actions_top)
        right.addLayout(actions_bottom)

        left_widget = QWidget()
        left_widget.setObjectName("sidebarCard")
        left_widget.setLayout(left)
        left_widget.setMinimumWidth(260)
        left_widget.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)

        right_widget = QWidget()
        right_widget.setObjectName("contentCard")
        right_widget.setLayout(right)
        right_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        splitter.setSizes([280, 900])

        root_layout = QHBoxLayout()
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.addWidget(splitter)
        wrapper = QWidget()
        wrapper.setLayout(root_layout)
        self.setCentralWidget(wrapper)

        self.apply_ui_preferences()
        self.refresh_events()

    def settings_store(self) -> QSettings:
        return QSettings(UI_SETTINGS_ORG, UI_SETTINGS_APP)

    def current_ui_preferences(self) -> dict[str, str | int]:
        return load_ui_preferences()

    def apply_ui_preferences(self) -> None:
        app = QApplication.instance()
        if app is None:
            return
        apply_ui_preferences(app)
        self.table.resizeColumnsToContents()

    def open_settings(self) -> None:
        dialog = AppearanceSettingsDialog(self.current_ui_preferences(), self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        selected = dialog.selected_preferences()
        settings = self.settings_store()
        settings.setValue("ui/theme", selected["theme"])
        settings.setValue("ui/font_family", selected["font_family"])
        settings.setValue("ui/font_size", selected["font_size"])
        self.apply_ui_preferences()
        QMessageBox.information(self, "Настройки", "Настройки интерфейса сохранены")

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
            item = QListWidgetItem(event.name)
            item.setData(Qt.ItemDataRole.UserRole, event.id)
            self.events_list.addItem(item)
        if self.events_list.count() > 0:
            self.events_list.setCurrentRow(0)

    def current_event_id(self) -> int | None:
        item = self.events_list.currentItem()
        if not item:
            return None
        event_id = item.data(Qt.ItemDataRole.UserRole)
        return int(event_id) if event_id is not None else None

    def load_swimmers(self) -> None:
        event_id = self.current_event_id()
        if event_id is None:
            self.table.setRowCount(0)
            return
        search_text = self.search_input.text().strip()
        search_all_events = bool(search_text)
        swimmers = self.service.repo.list_swimmers(None if search_all_events else event_id, search_text)
        self.table.clearSpans()
        self.table.setRowCount(len(swimmers))
        heat_rows: list[tuple[int | None, int]] = []
        for row_idx, s in enumerate(swimmers):
            values = [
                s.event_name or "",
                str(s.heat or "-"),
                str(s.lane or "-"),
                s.full_name,
                str(s.birth_year or ""),
                s.team or "",
                s.seed_time_raw or "",
                self._status_label(s.status),
                s.result_time_raw or "",
            ]
            for col_idx, val in enumerate(values):
                cell = QTableWidgetItem(val)
                if col_idx == 3:
                    cell.setData(Qt.ItemDataRole.UserRole, s.id)
                if col_idx in (1, 2):
                    cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if s.status == "DNS":
                    cell.setForeground(Qt.GlobalColor.darkGray)
                self.table.setItem(row_idx, col_idx, cell)
            heat_rows.append((s.heat, s.lane or 0))
        if search_all_events:
            self.table.clearSpans()
        else:
            self._apply_heat_spans(self.table, heat_rows)
        self.table.resizeColumnsToContents()

    def import_excel(self) -> None:
        settings = self.settings_store()
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
        ids: list[int] = []
        for idx in selected:
            item = self.table.item(idx.row(), 3)
            if item is None:
                continue
            swimmer_id = item.data(Qt.ItemDataRole.UserRole)
            if swimmer_id is not None:
                ids.append(int(swimmer_id))
        return ids

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
            build_html=lambda grouped, sort_by, sort_desc, group_by: self.service.build_event_protocol(
                event_id,
                grouped=grouped,
                sort_by=sort_by,
                sort_desc=sort_desc,
                group_by=group_by,
            ),
            build_excel=lambda path, grouped, sort_by, sort_desc, group_by: self.service.export_event_protocol_excel(
                path,
                event_id,
                grouped=grouped,
                sort_by=sort_by,
                sort_desc=sort_desc,
                group_by=group_by,
            ),
            allow_sorting=True,
            self_parent=self,
        )
        dialog.exec()

    def open_final_protocol(self) -> None:
        dialog = ProtocolDialog(
            self.service,
            title="Итоговый протокол соревнований",
            build_html=lambda grouped, sort_by, sort_desc, group_by: self.service.build_final_protocol(
                grouped=grouped,
                sort_by=sort_by,
                sort_desc=sort_desc,
                group_by=group_by,
            ),
            build_excel=lambda path, grouped, sort_by, sort_desc, group_by: self.service.export_final_protocol_excel(
                path,
                grouped=grouped,
                sort_by=sort_by,
                sort_desc=sort_desc,
                group_by=group_by,
            ),
            allow_sorting=True,
            self_parent=self,
        )
        dialog.exec()


class AppearanceSettingsDialog(QDialog):
    def __init__(self, preferences: dict[str, str | int], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Настройки интерфейса")
        self.resize(360, 180)

        form = QFormLayout()

        self.theme_combo = QComboBox()
        for theme_code, theme_label in THEME_LABELS.items():
            self.theme_combo.addItem(theme_label, theme_code)
        self.theme_combo.setCurrentIndex(self.theme_combo.findData(preferences["theme"]))

        self.font_combo = QComboBox()
        for family in FONT_FAMILIES:
            self.font_combo.addItem(family)
        self.font_combo.setCurrentText(str(preferences["font_family"]))

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 24)
        self.font_size_spin.setValue(int(preferences["font_size"]))

        form.addRow("Тема", self.theme_combo)
        form.addRow("Шрифт", self.font_combo)
        form.addRow("Размер шрифта", self.font_size_spin)

        buttons = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(save_btn)
        buttons.addWidget(cancel_btn)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addLayout(buttons)
        self.setLayout(layout)

    def selected_preferences(self) -> dict[str, str | int]:
        return {
            "theme": str(self.theme_combo.currentData()),
            "font_family": self.font_combo.currentText(),
            "font_size": self.font_size_spin.value(),
        }


class SecretaryAuthDialog(QDialog):
    def __init__(self, service: MeetService, parent: QWidget | None = None):
        super().__init__(parent)
        self.service = service
        self.authenticated_secretary: Secretary | None = None
        self.setWindowTitle("Авторизация секретаря соревнований")
        self.resize(520, 360)

        layout = QVBoxLayout()
        info = QLabel(
            "Зарегистрируйте любое количество секретарей. Для входа используйте логин и пароль, "
            "а при забытом пароле можно посмотреть сохранённую подсказку."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_login_tab(), "Авторизация")
        self.tabs.addTab(self._build_register_tab(), "Регистрация")
        self.tabs.addTab(self._build_recovery_tab(), "Забыл пароль")
        self.setLayout(layout)
        layout.addWidget(self.tabs)

    def _build_login_tab(self) -> QWidget:
        tab = QWidget()
        form = QFormLayout()
        self.login_username = QLineEdit()
        self.login_username.setPlaceholderText("Логин")
        self.login_password = QLineEdit()
        self.login_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.login_password.setPlaceholderText("Пароль")
        login_btn = QPushButton("Войти")
        login_btn.clicked.connect(self.handle_login)
        form.addRow("Логин", self.login_username)
        form.addRow("Пароль", self.login_password)
        form.addRow(login_btn)
        tab.setLayout(form)
        return tab

    def _build_register_tab(self) -> QWidget:
        tab = QWidget()
        form = QFormLayout()
        self.register_display_name = QLineEdit()
        self.register_display_name.setPlaceholderText("Например: Главный секретарь")
        self.register_username = QLineEdit()
        self.register_username.setPlaceholderText("Уникальный логин")
        self.register_password = QLineEdit()
        self.register_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.register_password.setPlaceholderText("Минимум 4 символа")
        self.register_hint = QLineEdit()
        self.register_hint.setPlaceholderText("Например: любимая команда")
        register_btn = QPushButton("Зарегистрировать")
        register_btn.clicked.connect(self.handle_register)
        form.addRow("Имя секретаря", self.register_display_name)
        form.addRow("Логин", self.register_username)
        form.addRow("Пароль", self.register_password)
        form.addRow("Подсказка", self.register_hint)
        form.addRow(register_btn)
        tab.setLayout(form)
        return tab

    def _build_recovery_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout()
        form = QFormLayout()
        self.recovery_username = QLineEdit()
        self.recovery_username.setPlaceholderText("Введите логин секретаря")
        show_hint_btn = QPushButton("Показать подсказку")
        show_hint_btn.clicked.connect(self.handle_show_hint)
        self.hint_output = QLabel("Подсказка появится здесь")
        self.hint_output.setWordWrap(True)
        form.addRow("Логин", self.recovery_username)
        form.addRow(show_hint_btn)
        layout.addLayout(form)
        layout.addWidget(self.hint_output)
        tab.setLayout(layout)
        return tab

    def handle_login(self) -> None:
        username = self.login_username.text().strip()
        password = self.login_password.text()
        secretary = self.service.authenticate_secretary(username, password)
        if secretary is None:
            QMessageBox.warning(self, "Авторизация", "Неверный логин или пароль")
            return
        self.authenticated_secretary = secretary
        self.accept()

    def handle_register(self) -> None:
        display_name = self.register_display_name.text().strip()
        username = self.register_username.text().strip()
        password = self.register_password.text()
        hint = self.register_hint.text().strip()
        try:
            self.service.register_secretary(username, password, hint, display_name=display_name)
        except ValueError as exc:
            QMessageBox.warning(self, "Регистрация", str(exc))
            return

        QMessageBox.information(
            self,
            "Регистрация",
            f"Секретарь {display_name or username} зарегистрирован. Теперь можно войти под этим логином.",
        )
        self.login_username.setText(username)
        self.login_password.clear()
        self.tabs.setCurrentIndex(0)
        self.register_display_name.clear()
        self.register_username.clear()
        self.register_password.clear()
        self.register_hint.clear()

    def handle_show_hint(self) -> None:
        username = self.recovery_username.text().strip()
        hint = self.service.get_secretary_password_hint(username)
        if hint is None:
            self.hint_output.setText("Секретарь с таким логином не найден.")
            return
        self.hint_output.setText(f"Подсказка для логина '{username}': {hint}")


class ResultsEntryDialog(QDialog):
    def __init__(self, service: MeetService, event_id: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.service = service
        self.event_id = event_id
        self.setWindowTitle("Ввод результатов заплыва")
        self.resize(900, 600)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Заплыв", "Дорожка", "ФИО", "Команда", "Заявка", "Результат", "Отметка"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setItemDelegateForColumn(5, TimeMaskDelegate(self.table))
        self.table.setItemDelegateForColumn(6, MarkDelegate(self.table))

        mark_hint = QLabel("Отметка: используйте коды судейства (например, DNS, DQ, EXH), если нужен комментарий к результату.")
        mark_hint.setWordWrap(True)

        save_btn = QPushButton("Сохранить результаты")
        save_btn.clicked.connect(self.save_results)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(mark_hint)
        layout.addWidget(save_btn)
        self.setLayout(layout)
        self.load_rows()

    def load_rows(self) -> None:
        swimmers = self.service.repo.list_swimmers(self.event_id)
        self.table.clearSpans()
        self.table.setRowCount(len(swimmers))
        heat_rows: list[tuple[int | None, int]] = []
        for row_idx, s in enumerate(swimmers):
            values = [
                str(s.heat or "-"),
                str(s.lane or "-"),
                s.full_name,
                s.team or "",
                s.seed_time_raw or "",
            ]
            for col_idx, val in enumerate(values):
                item = QTableWidgetItem(val)
                if col_idx == 2:
                    item.setData(Qt.ItemDataRole.UserRole, s.id)
                if col_idx < 5:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if col_idx in (0, 1):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(row_idx, col_idx, item)

            self.table.setItem(row_idx, 5, QTableWidgetItem(s.result_time_raw or ""))
            self.table.setItem(row_idx, 6, QTableWidgetItem(s.result_mark or ""))
            heat_rows.append((s.heat, s.lane or 0))
        MainWindow._apply_heat_spans(self.table, heat_rows)
        self.table.resizeColumnsToContents()

    def save_results(self) -> None:
        payload: list[dict[str, str]] = []
        for row in range(self.table.rowCount()):
            payload.append(
                {
                    "swimmer_id": str(self.table.item(row, 2).data(Qt.ItemDataRole.UserRole)),
                    "result_time_raw": self.table.item(row, 5).text() if self.table.item(row, 5) else "",
                    "result_mark": self.table.item(row, 6).text() if self.table.item(row, 6) else "",
                }
            )
        self.service.save_event_results(self.event_id, payload)
        QMessageBox.information(self, "Результаты", "Результаты сохранены")
        self.accept()


class TimeMaskDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setInputMask("00:00:00")
        return editor


class MarkDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setPlaceholderText("DNS / DQ / EXH")
        return editor


class ProtocolDialog(QDialog):
    def __init__(
        self,
        service: MeetService,
        title: str,
        build_html,
        build_excel=None,
        allow_sorting: bool = False,
        self_parent: QWidget | None = None,
    ):
        super().__init__(self_parent)
        self.service = service
        self.build_html = build_html
        self.build_excel = build_excel
        self.allow_sorting = allow_sorting
        self.setWindowTitle(title)
        self.resize(1000, 700)

        self.group_mode_combo = QComboBox()
        self.group_mode_combo.addItem("Группировка: по заплывам/дорожкам", "heat")
        if allow_sorting:
            self.group_mode_combo.addItem("Группировка: по команде", "team")
            self.group_mode_combo.addItem("Группировка: по году рождения", "birth_year")
            self.group_mode_combo.addItem("Группировка: по отметке", "mark")
            self.group_mode_combo.addItem("Группировка: по статусу", "status")
            self.group_mode_combo.addItem("Группировка: по дорожке", "lane")
        self.group_mode_combo.addItem("Группировка: без группировки", "none")
        self.group_mode_combo.currentIndexChanged.connect(self.refresh_html)

        self.sort_combo = QComboBox()
        self.sort_combo.addItem("Сортировка: по месту", "place")
        self.sort_combo.addItem("Сортировка: по ФИО", "full_name")
        self.sort_combo.addItem("Сортировка: по команде", "team")
        self.sort_combo.addItem("Сортировка: по году рождения", "birth_year")
        self.sort_combo.addItem("Сортировка: по заявке", "seed_time")
        self.sort_combo.addItem("Сортировка: по результату", "result_time")
        self.sort_combo.addItem("Сортировка: по заплыву", "heat")
        self.sort_combo.addItem("Сортировка: по дорожке", "lane")
        self.sort_combo.addItem("Сортировка: по статусу", "status")
        self.sort_combo.addItem("Сортировка: по отметке", "mark")
        self.sort_combo.currentIndexChanged.connect(self.refresh_html)
        self.sort_combo.setVisible(allow_sorting)

        self.place_sort_btn = QPushButton("Место ↑")
        self.place_sort_btn.setVisible(allow_sorting)
        self.place_sort_btn.clicked.connect(self.toggle_place_sort_order)
        self.sort_desc = False

        self.viewer = QTextEdit()
        self.viewer.setReadOnly(True)

        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_html)
        print_btn = QPushButton("Печать")
        print_btn.clicked.connect(self.print_protocol)
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_protocol)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self.group_mode_combo)
        toolbar.addWidget(self.sort_combo)
        toolbar.addWidget(self.place_sort_btn)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(print_btn)
        toolbar.addWidget(save_btn)

        layout = QVBoxLayout()
        layout.addLayout(toolbar)
        layout.addWidget(self.viewer)
        self.setLayout(layout)
        self.refresh_html()

    def current_html(self) -> str:
        group_mode = self.group_mode_combo.currentData()
        grouped = group_mode != "none"
        if self.allow_sorting:
            return self.build_html(grouped, self.sort_combo.currentData(), self.sort_desc, group_mode)
        return self.build_html(grouped)

    def toggle_place_sort_order(self) -> None:
        if not self.allow_sorting:
            return
        self.sort_combo.setCurrentIndex(self.sort_combo.findData("place"))
        self.sort_desc = not self.sort_desc
        self.place_sort_btn.setText("Место ↓" if self.sort_desc else "Место ↑")
        self.refresh_html()

    def refresh_html(self) -> None:
        self.viewer.setHtml(self.current_html())

    def _build_document(self) -> QTextDocument:
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setHtml(self.current_html())
        return doc

    def _configure_document_for_printer(self, doc: QTextDocument, printer: QPrinter) -> None:
        page_rect = printer.pageLayout().paintRect(QPageLayout.Unit.Point)
        doc.setPageSize(QSizeF(page_rect.width(), page_rect.height()))

    def print_protocol(self) -> None:
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            doc = self._build_document()
            self._configure_document_for_printer(doc, printer)
            doc.print_(printer)

    def save_protocol(self) -> None:
        path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Сохранить протокол",
            str(Path.home() / "protocol.pdf"),
            "PDF (*.pdf);;Excel (*.xlsx);;HTML (*.html);;Text (*.txt)",
        )
        if not path:
            return

        lower_path = path.lower()
        if selected_filter.startswith("PDF") or lower_path.endswith(".pdf"):
            if not lower_path.endswith(".pdf"):
                path = f"{path}.pdf"
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            doc = self._build_document()
            self._configure_document_for_printer(doc, printer)
            doc.print_(printer)
        elif selected_filter.startswith("Excel") or lower_path.endswith(".xlsx"):
            if self.build_excel is None:
                QMessageBox.warning(self, "Сохранение", "Экспорт в Excel для этого протокола недоступен")
                return
            if not lower_path.endswith(".xlsx"):
                path = f"{path}.xlsx"
            group_mode = self.group_mode_combo.currentData()
            grouped = group_mode != "none"
            if self.allow_sorting:
                self.build_excel(Path(path), grouped, self.sort_combo.currentData(), self.sort_desc, group_mode)
            else:
                self.build_excel(Path(path), grouped)
        else:
            text = self.current_html()
            Path(path).write_text(text, encoding="utf-8")
        QMessageBox.information(self, "Сохранение", f"Протокол сохранён: {path}")


def run_app() -> None:
    app = QApplication([])
    apply_ui_preferences(app)
    root = Path(__file__).resolve().parents[1]
    service = MeetService(root)
    auth_dialog = SecretaryAuthDialog(service)
    if auth_dialog.exec() != QDialog.DialogCode.Accepted or auth_dialog.authenticated_secretary is None:
        service.close()
        return
    window = MainWindow(service, root, auth_dialog.authenticated_secretary)
    window.show()
    app.exec()
    service.close()
