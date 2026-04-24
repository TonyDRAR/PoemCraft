from PySide6.QtWidgets import (
    QMainWindow, QTextEdit, QFileDialog, QMessageBox
)
from core.editor import Editor
from services.file_service import FileService


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Mon Éditeur")
        self.resize(900, 600)

        # Core
        self.editor_core = Editor()
        self.file_service = FileService()

        # UI
        self.text_edit = QTextEdit()
        self.setCentralWidget(self.text_edit)

        # Events
        self.text_edit.textChanged.connect(self.on_text_changed)

        self.create_menu()

    # ------------------------
    # UI
    # ------------------------
    def create_menu(self):
        menu = self.menuBar()

        file_menu = menu.addMenu("Fichier")

        new_action = file_menu.addAction("Nouveau")
        open_action = file_menu.addAction("Ouvrir")
        save_action = file_menu.addAction("Sauvegarder")
        save_as_action = file_menu.addAction("Sauvegarder sous")

        new_action.triggered.connect(self.new_file)
        open_action.triggered.connect(self.open_file)
        save_action.triggered.connect(self.save_file)
        save_as_action.triggered.connect(self.save_file_as)

    # ------------------------
    # Actions
    # ------------------------
    def new_file(self):
        if self.confirm_unsaved_changes():
            self.text_edit.clear()
            self.editor_core = Editor()

    def open_file(self):
        if not self.confirm_unsaved_changes():
            return

        path, _ = QFileDialog.getOpenFileName(self, "Ouvrir un fichier")

        if path:
            content = self.file_service.read(path)
            self.editor_core.set_content(content)
            self.editor_core.set_file_path(path)

            self.text_edit.setText(content)

    def save_file(self):
        if not self.editor_core.has_file():
            return self.save_file_as()

        content = self.text_edit.toPlainText()
        success = self.file_service.write(self.editor_core.file_path, content)

        if success:
            self.editor_core.set_content(content)

    def save_file_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Sauvegarder sous")

        if path:
            content = self.text_edit.toPlainText()
            success = self.file_service.write(path, content)

            if success:
                self.editor_core.set_file_path(path)
                self.editor_core.set_content(content)

    # ------------------------
    # Events
    # ------------------------
    def on_text_changed(self):
        self.editor_core.mark_modified()

    # ------------------------
    # Utils
    # ------------------------
    def confirm_unsaved_changes(self) -> bool:
        if not self.editor_core.is_modified():
            return True

        reply = QMessageBox.question(
            self,
            "Modifications non sauvegardées",
            "Voulez-vous sauvegarder avant de continuer ?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )

        if reply == QMessageBox.Yes:
            self.save_file()
            return True
        elif reply == QMessageBox.No:
            return True
        else:
            return False