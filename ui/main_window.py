import os
import tkinter as tk
from tkinter import filedialog, messagebox

from core.editor import Editor
from services.file_service import FileService


class MainWindow(tk.Tk):
    THEMES = {
        "light": {
            "window_bg": "#eef0f4",
            "surface_bg": "#ffffff",
            "surface_border": "#d7dbe3",
            "toolbar_bg": "#f8f9fb",
            "editor_bg": "#ffffff",
            "editor_fg": "#202124",
            "muted_fg": "#626a78",
            "insert_bg": "#202124",
            "select_bg": "#2f6fed",
            "select_fg": "#ffffff",
            "button_bg": "#ffffff",
            "button_fg": "#202124",
            "button_active_bg": "#e9eefb",
            "menu_bg": "#f8f9fb",
            "menu_fg": "#202124",
            "active_bg": "#e9eefb",
            "active_fg": "#202124",
        },
        "dark": {
            "window_bg": "#181a1f",
            "surface_bg": "#22252c",
            "surface_border": "#333844",
            "toolbar_bg": "#20232a",
            "editor_bg": "#15171c",
            "editor_fg": "#eceff4",
            "muted_fg": "#a4adba",
            "insert_bg": "#eceff4",
            "select_bg": "#4b7bec",
            "select_fg": "#ffffff",
            "button_bg": "#2b3039",
            "button_fg": "#eceff4",
            "button_active_bg": "#384152",
            "menu_bg": "#20232a",
            "menu_fg": "#eceff4",
            "active_bg": "#384152",
            "active_fg": "#ffffff",
        },
    }

    def __init__(self):
        super().__init__()

        self.title("Mon Editeur")
        self.geometry("980x680")
        self.minsize(720, 480)

        self.editor_core = Editor()
        self.file_service = FileService()
        self.dark_theme_enabled = tk.BooleanVar(value=False)
        self.menus = []
        self.toolbar_buttons = []

        self.create_widgets()
        self.create_menu()
        self.apply_theme()
        self.bind_shortcuts()
        self.update_status()
        self.protocol("WM_DELETE_WINDOW", self.close_window)

    def create_widgets(self):
        self.root_frame = tk.Frame(self, bd=0, highlightthickness=0)
        self.root_frame.pack(fill=tk.BOTH, expand=True)

        self.toolbar = tk.Frame(self.root_frame, bd=0, highlightthickness=0)
        self.toolbar.pack(fill=tk.X, padx=24, pady=(20, 0))

        self.title_block = tk.Frame(self.toolbar, bd=0, highlightthickness=0)
        self.title_block.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.app_title = tk.Label(
            self.title_block,
            text="Poetry Editor",
            anchor="w",
            font=("Segoe UI", 16, "bold"),
        )
        self.app_title.pack(anchor="w")

        self.file_label = tk.Label(
            self.title_block,
            text="Document sans titre",
            anchor="w",
            font=("Segoe UI", 9),
        )
        self.file_label.pack(anchor="w", pady=(2, 0))

        actions = [
            ("Nouveau", self.new_file),
            ("Ouvrir", self.open_file),
            ("Sauver", self.save_file),
        ]
        for label, command in actions:
            button = tk.Button(
                self.toolbar,
                text=label,
                command=command,
                bd=0,
                padx=14,
                pady=8,
                cursor="hand2",
                font=("Segoe UI", 9, "bold"),
            )
            button.pack(side=tk.LEFT, padx=(8, 0))
            self.toolbar_buttons.append(button)

        self.editor_shell = tk.Frame(self.root_frame, bd=0, highlightthickness=1)
        self.editor_shell.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

        self.editor_header = tk.Frame(self.editor_shell, bd=0, highlightthickness=0)
        self.editor_header.pack(fill=tk.X, padx=18, pady=(14, 8))

        self.mode_label = tk.Label(
            self.editor_header,
            text="Redaction",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        )
        self.mode_label.pack(side=tk.LEFT)

        self.hint_label = tk.Label(
            self.editor_header,
            text="Ctrl+S pour sauvegarder",
            anchor="e",
            font=("Segoe UI", 9),
        )
        self.hint_label.pack(side=tk.RIGHT)

        self.editor_body = tk.Frame(self.editor_shell, bd=0, highlightthickness=0)
        self.editor_body.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 16))

        self.scrollbar = tk.Scrollbar(self.editor_body, bd=0, width=14)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_edit = tk.Text(
            self.editor_body,
            wrap=tk.WORD,
            undo=True,
            bd=0,
            padx=26,
            pady=24,
            font=("Georgia", 14),
            spacing1=3,
            spacing2=2,
            spacing3=9,
            yscrollcommand=self.scrollbar.set,
        )
        self.text_edit.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.text_edit.yview)

        self.status_bar = tk.Frame(self.root_frame, bd=0, highlightthickness=0)
        self.status_bar.pack(fill=tk.X, padx=24, pady=(0, 18))

        self.status_left = tk.Label(self.status_bar, anchor="w", font=("Segoe UI", 9))
        self.status_left.pack(side=tk.LEFT)

        self.status_right = tk.Label(self.status_bar, anchor="e", font=("Segoe UI", 9))
        self.status_right.pack(side=tk.RIGHT)

        self.text_edit.bind("<<Modified>>", self.on_text_changed)
        self.text_edit.bind("<KeyRelease>", lambda _event: self.update_status())
        self.text_edit.bind("<ButtonRelease>", lambda _event: self.update_status())

    def create_menu(self):
        self.menu = tk.Menu(self)
        self.config(menu=self.menu)

        file_menu = tk.Menu(self.menu, tearoff=False)
        self.menu.add_cascade(label="Fichier", menu=file_menu)
        self.menus.append(file_menu)

        file_menu.add_command(label="Nouveau", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Ouvrir", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Sauvegarder", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Sauvegarder sous", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.close_window)

        options_menu = tk.Menu(self.menu, tearoff=False)
        self.menu.add_cascade(label="Options", menu=options_menu)
        self.menus.append(options_menu)

        options_menu.add_checkbutton(
            label="Theme sombre",
            variable=self.dark_theme_enabled,
            command=self.apply_theme,
        )

    def bind_shortcuts(self):
        self.bind("<Control-n>", lambda _event: self.new_file())
        self.bind("<Control-o>", lambda _event: self.open_file())
        self.bind("<Control-s>", lambda _event: self.save_file())

    def new_file(self):
        if self.confirm_unsaved_changes():
            self.text_edit.delete("1.0", tk.END)
            self.text_edit.edit_modified(False)
            self.editor_core = Editor()
            self.update_window_title()
            self.update_status()

    def open_file(self):
        if not self.confirm_unsaved_changes():
            return

        path = filedialog.askopenfilename(title="Ouvrir un fichier")

        if path:
            content = self.file_service.read(path)
            self.editor_core.set_content(content)
            self.editor_core.set_file_path(path)

            self.text_edit.delete("1.0", tk.END)
            self.text_edit.insert("1.0", content)
            self.text_edit.edit_modified(False)
            self.update_window_title()
            self.update_status()

    def save_file(self):
        if not self.editor_core.has_file():
            return self.save_file_as()

        content = self.get_text_content()
        success = self.file_service.write(self.editor_core.file_path, content)

        if success:
            self.editor_core.set_content(content)
            self.text_edit.edit_modified(False)
            self.update_window_title()
            self.update_status()

        return success

    def save_file_as(self):
        path = filedialog.asksaveasfilename(title="Sauvegarder sous")

        if not path:
            return False

        content = self.get_text_content()
        success = self.file_service.write(path, content)

        if success:
            self.editor_core.set_file_path(path)
            self.editor_core.set_content(content)
            self.text_edit.edit_modified(False)
            self.update_window_title()
            self.update_status()

        return success

    def on_text_changed(self, _event=None):
        if self.text_edit.edit_modified():
            self.editor_core.mark_modified()
            self.update_window_title()
            self.update_status()
            self.text_edit.edit_modified(False)

    def apply_theme(self):
        theme_name = "dark" if self.dark_theme_enabled.get() else "light"
        theme = self.THEMES[theme_name]

        self.configure(bg=theme["window_bg"])
        self.root_frame.configure(bg=theme["window_bg"])
        self.toolbar.configure(bg=theme["window_bg"])
        self.title_block.configure(bg=theme["window_bg"])
        self.app_title.configure(bg=theme["window_bg"], fg=theme["editor_fg"])
        self.file_label.configure(bg=theme["window_bg"], fg=theme["muted_fg"])

        self.editor_shell.configure(
            bg=theme["surface_bg"],
            highlightbackground=theme["surface_border"],
            highlightcolor=theme["surface_border"],
        )
        self.editor_header.configure(bg=theme["surface_bg"])
        self.mode_label.configure(bg=theme["surface_bg"], fg=theme["editor_fg"])
        self.hint_label.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.editor_body.configure(bg=theme["surface_bg"])
        self.status_bar.configure(bg=theme["window_bg"])
        self.status_left.configure(bg=theme["window_bg"], fg=theme["muted_fg"])
        self.status_right.configure(bg=theme["window_bg"], fg=theme["muted_fg"])

        self.scrollbar.configure(
            bg=theme["menu_bg"],
            activebackground=theme["active_bg"],
            troughcolor=theme["surface_bg"],
        )
        self.text_edit.configure(
            bg=theme["editor_bg"],
            fg=theme["editor_fg"],
            insertbackground=theme["insert_bg"],
            selectbackground=theme["select_bg"],
            selectforeground=theme["select_fg"],
        )
        self.menu.configure(
            bg=theme["menu_bg"],
            fg=theme["menu_fg"],
            activebackground=theme["active_bg"],
            activeforeground=theme["active_fg"],
        )

        for button in self.toolbar_buttons:
            button.configure(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["button_active_bg"],
                activeforeground=theme["button_fg"],
            )

        for menu in self.menus:
            menu.configure(
                bg=theme["menu_bg"],
                fg=theme["menu_fg"],
                activebackground=theme["active_bg"],
                activeforeground=theme["active_fg"],
                selectcolor=theme["editor_bg"],
            )

    def update_window_title(self):
        marker = "*" if self.editor_core.is_modified() else ""

        if self.editor_core.has_file():
            filename = os.path.basename(self.editor_core.file_path)
            self.title(f"{marker}{filename} - Mon Editeur")
            self.file_label.configure(text=self.editor_core.file_path)
            return

        self.title(f"{marker}Mon Editeur")
        self.file_label.configure(text="Document sans titre")

    def update_status(self):
        content = self.get_text_content()
        words = len(content.split())
        chars = len(content)
        line, column = self.text_edit.index(tk.INSERT).split(".")

        self.status_left.configure(text=f"{words} mots   {chars} caracteres")
        self.status_right.configure(text=f"Ligne {line}, colonne {int(column) + 1}")

    def confirm_unsaved_changes(self) -> bool:
        if not self.editor_core.is_modified():
            return True

        reply = messagebox.askyesnocancel(
            "Modifications non sauvegardees",
            "Voulez-vous sauvegarder avant de continuer ?",
            parent=self,
        )

        if reply is True:
            return bool(self.save_file())

        if reply is False:
            return True

        return False

    def close_window(self):
        if self.confirm_unsaved_changes():
            self.destroy()

    def get_text_content(self) -> str:
        return self.text_edit.get("1.0", "end-1c")
