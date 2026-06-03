import json
import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk

from PIL import Image, ImageTk

from core.editor import Editor
from services.file_service import FileService
from services.pollinations_service import PollinationsService


POEM_TYPES = [
    {
        "name": "Sonnet",
        "rules": "14 vers : 2 quatrains + 2 tercets. Souvent en alexandrins. Rimes ABBA ABBA CCD EDE ou variante proche.",
        "metric": "Alexandrin 12",
        "lines": [
            "[A] Premier vers du premier quatrain",
            "[B] Deuxieme vers du premier quatrain",
            "[B] Troisieme vers du premier quatrain",
            "[A] Quatrieme vers du premier quatrain",
            "",
            "[A] Premier vers du second quatrain",
            "[B] Deuxieme vers du second quatrain",
            "[B] Troisieme vers du second quatrain",
            "[A] Quatrieme vers du second quatrain",
            "",
            "[C] Premier vers du premier tercet",
            "[C] Deuxieme vers du premier tercet",
            "[D] Troisieme vers du premier tercet",
            "",
            "[E] Premier vers du second tercet",
            "[D] Deuxieme vers du second tercet",
            "[E] Troisieme vers du second tercet",
        ],
    },
    {
        "name": "Ballade",
        "rules": "3 strophes de meme longueur + un envoi final plus court. Meme refrain a la fin de chaque strophe.",
        "metric": "Octosyllabe 8",
        "lines": [
            "Strophe 1",
            "Vers 1",
            "Vers 2",
            "Vers 3",
            "{refrain_1}",
            "",
            "Strophe 2",
            "Vers 1",
            "Vers 2",
            "Vers 3",
            "{refrain_1}",
            "",
            "Strophe 3",
            "Vers 1",
            "Vers 2",
            "Vers 3",
            "{refrain_1}",
            "",
            "Envoi",
            "Vers 1",
            "Vers 2",
            "{refrain_1}",
        ],
    },
    {
        "name": "Rondeau",
        "rules": "Forme fixe, souvent 13 ou 15 vers, avec repetition partielle du premier vers comme refrain.",
        "metric": "Octosyllabe 8",
        "lines": [
            "Premier vers / refrain",
            "Vers 2",
            "Vers 3",
            "Vers 4",
            "Vers 5",
            "",
            "Refrain : {refrain_1}",
            "Vers 6",
            "Vers 7",
            "Vers 8",
            "",
            "Vers 9",
            "Vers 10",
            "Vers 11",
            "Vers 12",
            "Refrain : {refrain_1}",
        ],
    },
    {
        "name": "Villanelle",
        "rules": "19 vers : 5 tercets + 1 quatrain. Deux refrains alternes reviennent selon un schema precis.",
        "metric": "Libre",
        "lines": [
            "{refrain_1}",
            "Vers 2",
            "{refrain_2}",
            "",
            "Vers 4",
            "Vers 5",
            "{refrain_1}",
            "",
            "Vers 7",
            "Vers 8",
            "{refrain_2}",
            "",
            "Vers 10",
            "Vers 11",
            "{refrain_1}",
            "",
            "Vers 13",
            "Vers 14",
            "{refrain_2}",
            "",
            "Vers 16",
            "Vers 17",
            "{refrain_1}",
            "{refrain_2}",
        ],
    },
    {
        "name": "Ode",
        "rules": "Poeme lyrique celebrant une personne, une idee ou un evenement. Forme variable mais ton eleve.",
        "metric": "Libre",
        "lines": ["Strophe 1", "", "Strophe 2", "", "Strophe 3"],
    },
    {
        "name": "Elegie",
        "rules": "Poeme exprimant la tristesse, le regret ou le deuil. Forme libre.",
        "metric": "Libre",
        "lines": ["Strophe 1", "", "Strophe 2", "", "Strophe 3"],
    },
    {
        "name": "Fable",
        "rules": "Recit en vers ou en prose comportant une morale. Souvent avec des personnages personnifies.",
        "metric": "Libre",
        "lines": ["Situation", "", "Peripetie", "", "Denouement", "", "Morale : {morale}"],
    },
    {
        "name": "Poeme en vers libres",
        "rules": "Pas de nombre fixe de syllabes, ni de schema de rimes obligatoire. Grande liberte formelle.",
        "metric": "Libre",
        "lines": ["Vers 1", "Vers 2", "Vers 3"],
    },
    {
        "name": "Poeme en prose",
        "rules": "Ecrit en paragraphes et non en vers, avec images, rythme et musicalite.",
        "metric": "Libre",
        "lines": ["Premier paragraphe.", "", "Deuxieme paragraphe."],
    },
    {
        "name": "Haiku",
        "rules": "3 vers. Traditionnellement 5-7-5 syllabes. Evoque souvent la nature et un instant fugace.",
        "metric": "Haiku 5/7/5",
        "lines": ["[5] Premier vers", "[7] Deuxieme vers", "[5] Troisieme vers"],
    },
    {
        "name": "Acrostiche",
        "rules": "Les premieres lettres des vers forment un mot ou une phrase lorsqu'on les lit verticalement.",
        "metric": "Libre",
        "lines": [],
    },
]


class PoemCreationDialog(tk.Toplevel):
    def __init__(self, parent, poem_types: list[dict], theme: dict):
        super().__init__(parent)
        self.parent = parent
        self.poem_types = poem_types
        self.theme = theme
        self.result = None
        self.title("Ajouter un poeme")
        self.geometry("760x520")
        self.minsize(640, 440)
        self.transient(parent)
        self.grab_set()

        self.title_var = tk.StringVar(value="Nouveau poeme")
        self.refrain_1_var = tk.StringVar(value="Refrain principal")
        self.refrain_2_var = tk.StringVar(value="Second refrain")
        self.acrostic_var = tk.StringVar(value="POEME")
        self.morale_var = tk.StringVar(value="Morale a formuler")
        self.selected_index = 0
        self.type_buttons = []

        self.create_widgets()
        for variable in (
            self.title_var,
            self.refrain_1_var,
            self.refrain_2_var,
            self.acrostic_var,
            self.morale_var,
        ):
            variable.trace_add("write", self.update_preview)
        self.apply_theme()
        self.select_type(0)
        self.bind("<Escape>", lambda _event: self.cancel())
        self.bind("<Return>", lambda _event: self.create_poem())
        self.wait_window()

    def create_widgets(self):
        self.container = tk.Frame(self, bd=0, highlightthickness=0)
        self.container.pack(fill=tk.BOTH, expand=True, padx=22, pady=20)

        self.header = tk.Frame(self.container, bd=0, highlightthickness=0)
        self.header.pack(fill=tk.X)

        self.heading = tk.Label(self.header, text="Ajouter un poeme", anchor="w", font=("Segoe UI", 16, "bold"))
        self.heading.pack(anchor="w")

        self.subtitle = tk.Label(
            self.header,
            text="Choisissez une forme, renseignez les champs utiles, puis creez un brouillon structure.",
            anchor="w",
            font=("Segoe UI", 9),
        )
        self.subtitle.pack(anchor="w", pady=(4, 0))

        self.body = tk.Frame(self.container, bd=0, highlightthickness=0)
        self.body.pack(fill=tk.BOTH, expand=True, pady=(18, 0))

        self.type_list = tk.Frame(self.body, bd=0, highlightthickness=1, width=230)
        self.type_list.pack(side=tk.LEFT, fill=tk.Y)
        self.type_list.pack_propagate(False)

        for index, poem_type in enumerate(self.poem_types):
            button = tk.Button(
                self.type_list,
                text=poem_type["name"],
                command=lambda selected=index: self.select_type(selected),
                anchor="w",
                bd=0,
                padx=12,
                pady=8,
                cursor="hand2",
                font=("Segoe UI", 9, "bold"),
            )
            button.pack(fill=tk.X)
            self.type_buttons.append(button)

        self.details = tk.Frame(self.body, bd=0, highlightthickness=1)
        self.details.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 0))

        self.form = tk.Frame(self.details, bd=0, highlightthickness=0)
        self.form.pack(fill=tk.X, padx=18, pady=18)

        self.title_label = tk.Label(self.form, text="Titre", anchor="w", font=("Segoe UI", 9, "bold"))
        self.title_label.pack(anchor="w")
        self.title_entry = tk.Entry(self.form, textvariable=self.title_var, bd=0, font=("Segoe UI", 10))
        self.title_entry.pack(fill=tk.X, pady=(5, 12), ipady=8)

        self.refrain_1_label = tk.Label(self.form, text="Refrain principal", anchor="w", font=("Segoe UI", 9, "bold"))
        self.refrain_1_entry = tk.Entry(self.form, textvariable=self.refrain_1_var, bd=0, font=("Segoe UI", 10))
        self.refrain_2_label = tk.Label(self.form, text="Second refrain", anchor="w", font=("Segoe UI", 9, "bold"))
        self.refrain_2_entry = tk.Entry(self.form, textvariable=self.refrain_2_var, bd=0, font=("Segoe UI", 10))
        self.acrostic_label = tk.Label(self.form, text="Mot de l'acrostiche", anchor="w", font=("Segoe UI", 9, "bold"))
        self.acrostic_entry = tk.Entry(self.form, textvariable=self.acrostic_var, bd=0, font=("Segoe UI", 10))
        self.morale_label = tk.Label(self.form, text="Morale", anchor="w", font=("Segoe UI", 9, "bold"))
        self.morale_entry = tk.Entry(self.form, textvariable=self.morale_var, bd=0, font=("Segoe UI", 10))

        self.rules_label = tk.Label(self.details, text="Regles principales", anchor="w", font=("Segoe UI", 10, "bold"))
        self.rules_label.pack(anchor="w", padx=18, pady=(6, 0))

        self.rules_text = tk.Message(self.details, text="", anchor="nw", font=("Segoe UI", 10), width=420)
        self.rules_text.pack(fill=tk.X, padx=18, pady=(8, 0))

        self.preview_label = tk.Label(self.details, text="Structure generee", anchor="w", font=("Segoe UI", 10, "bold"))
        self.preview_label.pack(anchor="w", padx=18, pady=(18, 0))

        self.preview = tk.Text(self.details, height=8, bd=0, padx=12, pady=10, font=("Consolas", 9), wrap=tk.WORD)
        self.preview.pack(fill=tk.BOTH, expand=True, padx=18, pady=(8, 18))
        self.preview.configure(state=tk.DISABLED)

        self.footer = tk.Frame(self.container, bd=0, highlightthickness=0)
        self.footer.pack(fill=tk.X, pady=(18, 0))

        self.cancel_button = tk.Button(
            self.footer,
            text="Annuler",
            command=self.cancel,
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        )
        self.cancel_button.pack(side=tk.RIGHT)

        self.create_button = tk.Button(
            self.footer,
            text="Ajouter le poeme",
            command=self.create_poem,
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        )
        self.create_button.pack(side=tk.RIGHT, padx=(0, 8))
        self.body.pack_forget()
        self.footer.pack_forget()
        self.footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(14, 0))
        self.body.pack(fill=tk.BOTH, expand=True, pady=(18, 0))

    def apply_theme(self):
        theme = self.theme
        self.configure(bg=theme["window_bg"])

        for frame in (self.container, self.header, self.body, self.footer):
            frame.configure(bg=theme["window_bg"])

        for frame in (self.type_list, self.details):
            frame.configure(
                bg=theme["surface_bg"],
                highlightbackground=theme["surface_border"],
                highlightcolor=theme["surface_border"],
            )

        self.form.configure(bg=theme["surface_bg"])

        labels = (
            self.heading,
            self.title_label,
            self.refrain_1_label,
            self.refrain_2_label,
            self.acrostic_label,
            self.morale_label,
            self.rules_label,
            self.preview_label,
        )
        for label in labels:
            label.configure(bg=theme["window_bg"] if label in (self.heading,) else theme["surface_bg"], fg=theme["editor_fg"])

        self.subtitle.configure(bg=theme["window_bg"], fg=theme["muted_fg"])
        self.rules_text.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.preview.configure(bg=theme["editor_bg"], fg=theme["editor_fg"], insertbackground=theme["insert_bg"])

        for entry in (
            self.title_entry,
            self.refrain_1_entry,
            self.refrain_2_entry,
            self.acrostic_entry,
            self.morale_entry,
        ):
            entry.configure(
                bg=theme["editor_bg"],
                fg=theme["editor_fg"],
                insertbackground=theme["insert_bg"],
                highlightthickness=1,
                highlightbackground=theme["surface_border"],
                highlightcolor=theme["select_bg"],
            )

        for button in self.type_buttons + [self.cancel_button, self.create_button]:
            button.configure(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["button_active_bg"],
                activeforeground=theme["button_fg"],
            )

        self.create_button.configure(
            bg=theme["select_bg"],
            fg=theme["select_fg"],
            activebackground=theme["select_bg"],
            activeforeground=theme["select_fg"],
        )

    def select_type(self, index: int):
        self.selected_index = index
        poem_type = self.poem_types[index]
        theme = self.theme

        for button_index, button in enumerate(self.type_buttons):
            is_active = button_index == index
            button.configure(
                bg=theme["select_bg"] if is_active else theme["button_bg"],
                fg=theme["select_fg"] if is_active else theme["button_fg"],
            )

        self.rules_text.configure(text=poem_type["rules"])
        self.create_button.configure(text=f"Ajouter : {poem_type['name']}")
        self.update_field_visibility(poem_type["name"])
        self.update_preview()

    def update_field_visibility(self, poem_name: str):
        optional_fields = (
            self.refrain_1_label,
            self.refrain_1_entry,
            self.refrain_2_label,
            self.refrain_2_entry,
            self.acrostic_label,
            self.acrostic_entry,
            self.morale_label,
            self.morale_entry,
        )

        for widget in optional_fields:
            widget.pack_forget()

        if poem_name in ("Ballade", "Rondeau", "Villanelle"):
            self.refrain_1_label.pack(anchor="w")
            self.refrain_1_entry.pack(fill=tk.X, pady=(5, 12), ipady=8)

        if poem_name == "Villanelle":
            self.refrain_2_label.pack(anchor="w")
            self.refrain_2_entry.pack(fill=tk.X, pady=(5, 12), ipady=8)

        if poem_name == "Acrostiche":
            self.acrostic_label.pack(anchor="w")
            self.acrostic_entry.pack(fill=tk.X, pady=(5, 12), ipady=8)

        if poem_name == "Fable":
            self.morale_label.pack(anchor="w")
            self.morale_entry.pack(fill=tk.X, pady=(5, 12), ipady=8)

    def build_preview_content(self) -> str:
        poem_type = self.poem_types[self.selected_index]
        return build_poem_draft(
            poem_type,
            self.title_var.get(),
            self.refrain_1_var.get(),
            self.refrain_2_var.get(),
            self.acrostic_var.get(),
            self.morale_var.get(),
        )

    def update_preview(self, *_args):
        content = self.build_preview_content()
        self.preview.configure(state=tk.NORMAL)
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", content)
        self.preview.configure(state=tk.DISABLED)

    def create_poem(self):
        self.result = {
            "poem_type": self.poem_types[self.selected_index],
            "title": self.title_var.get().strip() or self.poem_types[self.selected_index]["name"],
            "content": self.build_preview_content(),
        }
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


def build_poem_draft(poem_type: dict, title: str, refrain_1: str, refrain_2: str, acrostic: str, morale: str) -> str:
    title = title.strip() or poem_type["name"]
    placeholders = {
        "refrain_1": refrain_1.strip() or "Refrain principal",
        "refrain_2": refrain_2.strip() or "Second refrain",
        "morale": morale.strip() or "Morale a formuler",
    }

    if poem_type["name"] == "Acrostiche":
        letters = [character.upper() for character in acrostic.strip() if character.strip()]
        if not letters:
            letters = list("POEME")
        lines = [f"{letter}..." for letter in letters]
    else:
        lines = [line.format(**placeholders) for line in poem_type["lines"]]

    return "\n".join([title, poem_type["name"], "", *lines]).rstrip() + "\n"


def get_poem_type_by_name(name: str) -> dict | None:
    normalized_name = name.strip().lower()

    for poem_type in POEM_TYPES:
        if poem_type["name"].lower() == normalized_name:
            return poem_type

    return None


def detect_poem_type_from_content(content: str) -> dict | None:
    lines = content.splitlines()

    if len(lines) < 2:
        return None

    return get_poem_type_by_name(lines[1])


def is_structure_heading(line: str) -> bool:
    normalized_line = line.strip().lower()
    return normalized_line in {"situation", "peripetie", "denouement", "envoi"} or normalized_line.startswith("strophe ")


def get_poem_body_lines(content: str, poem_type: dict) -> list[str]:
    lines = content.splitlines()

    if len(lines) >= 2 and lines[1].strip().lower() == poem_type["name"].lower():
        return lines[3:] if len(lines) >= 3 and not lines[2].strip() else lines[2:]

    return lines


def build_editor_structure_text(poem_type: dict, content: str) -> str:
    body_lines = get_poem_body_lines(content, poem_type)
    display_lines = []
    verse_index = 0

    for line in body_lines:
        stripped_line = line.strip()

        if not stripped_line:
            if display_lines and display_lines[-1]:
                display_lines.append("")
            continue

        if is_structure_heading(stripped_line):
            display_lines.append(stripped_line.upper())
            continue

        verse_index += 1
        display_lines.append(f"{verse_index:02d}  {stripped_line}")

    if not display_lines:
        display_lines.append("La structure apparaitra ici des que le poeme contient des lignes.")

    return "\n".join(display_lines).strip()


def slugify_filename(name: str) -> str:
    normalized = "".join(character if character.isalnum() else "_" for character in name.strip().lower())
    normalized = "_".join(part for part in normalized.split("_") if part)
    return normalized or "nouveau_poeme"


class MainWindow(tk.Tk):
    THEMES = {
        "light": {
            "window_bg": "#eef0f4",
            "chrome_bg": "#f7f8fb",
            "chrome_border": "#dfe3eb",
            "chrome_mark_bg": "#202124",
            "chrome_mark_fg": "#ffffff",
            "chrome_button_hover": "#e8edf6",
            "chrome_close_hover": "#d94b4b",
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
            "tree_selected_bg": "#dbe7ff",
            "scrollbar_bg": "#d7dbe3",
            "scrollbar_active_bg": "#b8bfcc",
            "metric_ok_fg": "#1f8a4c",
            "metric_error_fg": "#d14a32",
            "rhyme_palette": ("#2f6fed", "#b2569d", "#c67a00", "#1f8a4c", "#7c5fb8", "#168a8a"),
        },
        "dark": {
            "window_bg": "#181a1f",
            "chrome_bg": "#111318",
            "chrome_border": "#2c313b",
            "chrome_mark_bg": "#eceff4",
            "chrome_mark_fg": "#15171c",
            "chrome_button_hover": "#2b3039",
            "chrome_close_hover": "#b84040",
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
            "tree_selected_bg": "#334467",
            "scrollbar_bg": "#3a404c",
            "scrollbar_active_bg": "#505a6b",
            "metric_ok_fg": "#74c98f",
            "metric_error_fg": "#ff8a74",
            "rhyme_palette": ("#82aaff", "#f29fca", "#ffc46b", "#74c98f", "#b7a4ff", "#6fd6d6"),
        },
    }
    TEXT_FILETYPES = (
        ("Fichiers texte", "*.txt"),
        ("Tous les fichiers", "*.*"),
    )
    IMAGE_METADATA_FILENAME = ".poetry_editor_images.json"
    IMAGE_ASSETS_FOLDER = ".poetry_editor_assets"
    SETTINGS_FILENAME = "settings.json"

    def __init__(self):
        super().__init__()

        self.title("Mon Editeur")
        self.geometry("980x680")
        self.minsize(720, 480)
        self.overrideredirect(True)

        self.editor_core = Editor()
        self.file_service = FileService()
        self.pollinations_service = PollinationsService()
        self.app_settings = self.load_app_settings()
        self.dark_theme_enabled = tk.BooleanVar(value=bool(self.app_settings.get("dark_theme_enabled", False)))
        self.metric_objective = tk.StringVar(value=self.get_saved_metric_objective())
        self.analysis_mode = tk.StringVar(value="")
        self.menus = []
        self.toolbar_buttons = []
        self.window_control_buttons = []
        self.resize_grips = []
        self.syllable_line_counts = []
        self.syllable_line_targets = []
        self.rhyme_line_labels = []
        self.syllable_count_pending = False
        self.current_folder = None
        self.current_image_path = None
        self.current_poem_type = None
        self.image_preview = None
        self.image_generation_pending = False
        self.image_generation_id = 0
        self.generate_image_button = None
        self.is_maximized = False
        self.restore_geometry = ""
        self.drag_offset_x = 0
        self.drag_offset_y = 0
        self.resize_start_width = 0
        self.resize_start_height = 0
        self.resize_start_x = 0
        self.resize_start_y = 0
        self.resize_start_window_x = 0
        self.resize_start_window_y = 0
        self.resize_direction = ""
        self.resize_margin = 7
        self.folder_tree_style_name = "Poetry.Treeview"
        self.scrollbar_style_name = "Poetry.Vertical.TScrollbar"
        self.metric_selector_style_name = "Poetry.TCombobox"
        self.ui_style = ttk.Style(self)

        try:
            self.ui_style.theme_use("clam")
        except tk.TclError:
            pass

        self.create_widgets()
        self.create_context_menus()
        self.apply_theme()
        self.bind_shortcuts()
        self.restore_session()
        self.update_status()
        self.bind("<Map>", self.restore_custom_window_chrome)
        self.protocol("WM_DELETE_WINDOW", self.close_window)

    def create_widgets(self):
        self.root_frame = tk.Frame(self, bd=0, highlightthickness=0)
        self.root_frame.pack(fill=tk.BOTH, expand=True)

        self.window_chrome = tk.Frame(self.root_frame, bd=0, highlightthickness=0, height=46)
        self.window_chrome.pack(fill=tk.X)
        self.window_chrome.pack_propagate(False)

        self.window_title_area = tk.Frame(self.window_chrome, bd=0, highlightthickness=0)
        self.window_title_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(18, 8))

        self.app_mark = tk.Label(
            self.window_title_area,
            text="P",
            anchor="center",
            font=("Segoe UI", 10, "bold"),
            width=3,
        )
        self.app_mark.pack(side=tk.LEFT, pady=10)

        self.window_text_block = tk.Frame(self.window_title_area, bd=0, highlightthickness=0)
        self.window_text_block.pack(side=tk.LEFT, padx=(10, 0), pady=13)

        self.window_app_name = tk.Label(
            self.window_text_block,
            text="PoemCraft",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        )
        self.window_app_name.pack(anchor="w")

        self.window_actions = tk.Frame(self.window_chrome, bd=0, highlightthickness=0)
        self.window_actions.pack(side=tk.RIGHT, fill=tk.Y)

        window_actions = [
            ("-", self.minimize_window),
            ("□", self.toggle_maximize_window),
            ("×", self.close_window),
        ]

        for label, command in window_actions:
            button = tk.Button(
                self.window_actions,
                text=label,
                command=command,
                bd=0,
                width=5,
                cursor="hand2",
                font=("Segoe UI", 11, "bold"),
            )
            button.pack(side=tk.LEFT, fill=tk.Y)
            self.window_control_buttons.append(button)

        for widget in (
            self.window_chrome,
            self.window_title_area,
            self.app_mark,
            self.window_text_block,
            self.window_app_name,
        ):
            widget.bind("<Button-1>", self.start_window_drag)
            widget.bind("<B1-Motion>", self.drag_window)
            widget.bind("<Double-Button-1>", lambda _event: self.toggle_maximize_window())

        self.toolbar = tk.Frame(self.root_frame, bd=0, highlightthickness=0)
        self.toolbar.pack(fill=tk.X, padx=24, pady=(14, 0))

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
            ("Ajouter un poeme", self.add_new_poem),
            ("Nouveau", self.new_file),
            ("Ouvrir", self.open_file),
            ("Sauver", self.save_file),
            ("Theme sombre", self.toggle_theme),
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

        self.workspace = tk.PanedWindow(
            self.root_frame,
            orient=tk.HORIZONTAL,
            bd=0,
            sashwidth=6,
            showhandle=False,
        )
        self.workspace.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

        self.sidebar_shell = tk.Frame(self.workspace, bd=0, highlightthickness=1)
        self.workspace.add(self.sidebar_shell, minsize=180, width=240)

        self.sidebar_header = tk.Frame(self.sidebar_shell, bd=0, highlightthickness=0)
        self.sidebar_header.pack(fill=tk.X, padx=12, pady=(12, 8))

        self.sidebar_title = tk.Label(
            self.sidebar_header,
            text="Explorateur",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        )
        self.sidebar_title.pack(side=tk.LEFT)

        self.folder_button = tk.Button(
            self.sidebar_header,
            text="Ouvrir",
            command=self.open_folder,
            bd=0,
            padx=10,
            pady=5,
            cursor="hand2",
            font=("Segoe UI", 8, "bold"),
        )
        self.folder_button.pack(side=tk.RIGHT)
        self.toolbar_buttons.append(self.folder_button)

        self.folder_label = tk.Label(
            self.sidebar_shell,
            text="Aucun dossier ouvert",
            anchor="w",
            font=("Segoe UI", 8),
        )
        self.folder_label.pack(fill=tk.X, padx=12, pady=(0, 8))

        self.tree_frame = tk.Frame(self.sidebar_shell, bd=0, highlightthickness=0)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.folder_tree = ttk.Treeview(
            self.tree_frame,
            columns=("path",),
            displaycolumns=(),
            show="tree",
            selectmode="browse",
            style=self.folder_tree_style_name,
        )
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_scrollbar = ttk.Scrollbar(
            self.tree_frame,
            orient=tk.VERTICAL,
            style=self.scrollbar_style_name,
        )
        self.tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.configure(yscrollcommand=self.tree_scrollbar.set)
        self.tree_scrollbar.configure(command=self.folder_tree.yview)
        self.folder_tree.bind("<<TreeviewOpen>>", self.on_tree_open)
        self.folder_tree.bind("<Double-1>", self.open_selected_tree_file)
        self.folder_tree.bind("<Button-3>", self.show_explorer_context_menu)
        self.folder_tree.bind("<Delete>", lambda _event: self.delete_selected_explorer_item())

        self.editor_shell = tk.Frame(self.workspace, bd=0, highlightthickness=1)
        self.workspace.add(self.editor_shell, minsize=360)

        self.editor_header = tk.Frame(self.editor_shell, bd=0, highlightthickness=0)
        self.editor_header.pack(fill=tk.X, padx=18, pady=(14, 8))

        self.editor_title_block = tk.Frame(self.editor_header, bd=0, highlightthickness=0)
        self.editor_title_block.pack(side=tk.LEFT)

        self.mode_label = tk.Label(
            self.editor_title_block,
            text="Redaction",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        )
        self.mode_label.pack(anchor="w")

        self.analysis_bar = tk.Frame(self.editor_title_block, bd=0, highlightthickness=0)
        self.analysis_bar.pack(anchor="w", pady=(8, 0))

        self.analysis_label = tk.Label(
            self.analysis_bar,
            text="Analyse",
            anchor="w",
            font=("Segoe UI", 8),
        )
        self.analysis_label.pack(side=tk.LEFT, padx=(0, 8))

        self.syllables_button = tk.Button(
            self.analysis_bar,
            text="Syllabes",
            command=self.show_syllable_count,
            bd=0,
            padx=10,
            pady=6,
            cursor="hand2",
            font=("Segoe UI", 8, "bold"),
        )
        self.syllables_button.pack(side=tk.LEFT)
        self.toolbar_buttons.append(self.syllables_button)

        self.rhymes_button = tk.Button(
            self.analysis_bar,
            text="Rimes",
            command=self.show_rhyme_scheme,
            bd=0,
            padx=10,
            pady=6,
            cursor="hand2",
            font=("Segoe UI", 8, "bold"),
        )
        self.rhymes_button.pack(side=tk.LEFT, padx=(6, 0))
        self.toolbar_buttons.append(self.rhymes_button)

        self.metric_frame = tk.Frame(self.analysis_bar, bd=0, highlightthickness=0)
        self.metric_frame.pack(side=tk.LEFT, padx=(14, 0))

        self.metric_label = tk.Label(
            self.metric_frame,
            text="Objectif",
            anchor="w",
            font=("Segoe UI", 8),
        )
        self.metric_label.pack(side=tk.LEFT, padx=(0, 6))

        self.metric_selector = ttk.Combobox(
            self.metric_frame,
            textvariable=self.metric_objective,
            values=Editor.get_metric_names(),
            state="readonly",
            style=self.metric_selector_style_name,
            width=17,
        )
        self.metric_selector.pack(side=tk.LEFT)
        self.metric_selector.bind("<<ComboboxSelected>>", self.on_metric_objective_changed)

        self.editor_tools = tk.Frame(self.editor_header, bd=0, highlightthickness=0)
        self.editor_tools.pack(side=tk.RIGHT, padx=(12, 0))

        editor_actions = [
            ("Image IA", self.import_image_for_current_file),
        ]

        for label, command in editor_actions:
            button = tk.Button(
                self.editor_tools,
                text=label,
                command=command,
                bd=0,
                padx=10,
                pady=6,
                cursor="hand2",
                font=("Segoe UI", 8, "bold"),
            )
            button.pack(side=tk.LEFT, padx=(0, 6))
            self.toolbar_buttons.append(button)
            if label == "Image IA":
                self.generate_image_button = button

        self.hint_label = tk.Label(
            self.editor_header,
            text="Ctrl+S pour sauvegarder",
            anchor="e",
            font=("Segoe UI", 9),
        )
        self.hint_label.pack(side=tk.RIGHT)

        self.editor_content = tk.Frame(self.editor_shell, bd=0, highlightthickness=0)
        self.editor_content.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 16))

        self.editor_body = tk.Frame(self.editor_content, bd=0, highlightthickness=0)
        self.editor_body.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.structure_panel = tk.Frame(self.editor_content, bd=0, highlightthickness=1, width=280)
        self.structure_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(18, 0))
        self.structure_panel.pack_propagate(False)

        self.structure_header = tk.Frame(self.structure_panel, bd=0, highlightthickness=0)
        self.structure_header.pack(fill=tk.X, padx=14, pady=(14, 6))

        self.structure_title = tk.Label(
            self.structure_header,
            text="Structure",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        )
        self.structure_title.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.hide_structure_button = tk.Button(
            self.structure_header,
            text="Masquer",
            command=self.hide_poem_structure,
            bd=0,
            padx=8,
            pady=5,
            cursor="hand2",
            font=("Segoe UI", 8, "bold"),
        )
        self.hide_structure_button.pack(side=tk.RIGHT)
        self.toolbar_buttons.append(self.hide_structure_button)

        self.structure_type_label = tk.Label(
            self.structure_panel,
            text="",
            anchor="w",
            font=("Segoe UI", 12, "bold"),
        )
        self.structure_type_label.pack(fill=tk.X, padx=14, pady=(2, 0))

        self.structure_metric_label = tk.Label(
            self.structure_panel,
            text="",
            anchor="w",
            font=("Segoe UI", 9),
        )
        self.structure_metric_label.pack(fill=tk.X, padx=14, pady=(4, 0))

        self.structure_rules_text = tk.Message(
            self.structure_panel,
            text="",
            anchor="nw",
            font=("Segoe UI", 9),
            width=245,
        )
        self.structure_rules_text.pack(fill=tk.X, padx=14, pady=(12, 0))

        self.structure_list_label = tk.Label(
            self.structure_panel,
            text="Plan du brouillon",
            anchor="w",
            font=("Segoe UI", 9, "bold"),
        )
        self.structure_list_label.pack(fill=tk.X, padx=14, pady=(16, 6))

        self.structure_text_frame = tk.Frame(self.structure_panel, bd=0, highlightthickness=0)
        self.structure_text_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

        self.structure_text = tk.Text(
            self.structure_text_frame,
            bd=0,
            padx=10,
            pady=10,
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self.structure_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.structure_text.configure(state=tk.DISABLED)

        self.structure_scrollbar = ttk.Scrollbar(
            self.structure_text_frame,
            orient=tk.VERTICAL,
            style=self.scrollbar_style_name,
        )
        self.structure_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.structure_text.configure(yscrollcommand=self.structure_scrollbar.set)
        self.structure_scrollbar.configure(command=self.structure_text.yview)
        self.structure_panel.pack_forget()

        self.image_panel = tk.Frame(self.editor_content, bd=0, highlightthickness=0, width=340)
        self.image_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(18, 0))
        self.image_panel.pack_propagate(False)

        self.image_preview_label = tk.Label(
            self.image_panel,
            anchor="n",
            bd=0,
            padx=0,
            pady=0,
        )
        self.image_preview_label.pack(fill=tk.X)

        self.remove_image_button = tk.Button(
            self.image_panel,
            text="Retirer l'image",
            command=self.remove_image_from_current_file,
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
            font=("Segoe UI", 8, "bold"),
        )
        self.remove_image_button.pack(anchor="w", pady=(12, 0))
        self.toolbar_buttons.append(self.remove_image_button)
        self.image_panel.pack_forget()

        self.scrollbar = ttk.Scrollbar(
            self.editor_body,
            orient=tk.VERTICAL,
            style=self.scrollbar_style_name,
        )
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.syllable_gutter = tk.Canvas(
            self.editor_body,
            width=38,
            bd=0,
            highlightthickness=0,
        )
        self.syllable_gutter.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        self.text_edit = tk.Text(
            self.editor_body,
            wrap=tk.WORD,
            undo=True,
            bd=0,
            padx=20,
            pady=24,
            font=("Georgia", 14),
            spacing1=3,
            spacing2=2,
            spacing3=9,
            yscrollcommand=self.on_text_scroll,
        )
        self.text_edit.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.scroll_text)

        self.status_bar = tk.Frame(self.root_frame, bd=0, highlightthickness=0)
        self.status_bar.pack(fill=tk.X, padx=24, pady=(0, 18))

        self.status_left = tk.Label(self.status_bar, anchor="w", font=("Segoe UI", 9))
        self.status_left.pack(side=tk.LEFT)

        self.status_right = tk.Label(self.status_bar, anchor="e", font=("Segoe UI", 9))
        self.status_right.pack(side=tk.RIGHT)

        self.text_edit.bind("<<Modified>>", self.on_text_changed)
        self.text_edit.bind("<KeyRelease>", self.on_editor_navigation)
        self.text_edit.bind("<ButtonRelease>", self.on_editor_navigation)
        self.text_edit.bind("<Configure>", lambda _event: self.schedule_syllable_gutter_redraw())
        self.text_edit.bind("<MouseWheel>", lambda _event: self.schedule_syllable_gutter_redraw(), add="+")
        self.create_resize_grips()

    def create_context_menus(self):
        self.explorer_context_menu = tk.Menu(self, tearoff=False)
        self.menus.append(self.explorer_context_menu)
        self.explorer_context_menu.add_command(label="Ajouter un poeme", command=self.create_poem_from_explorer)
        self.explorer_context_menu.add_command(label="Nouveau texte", command=self.create_text_from_explorer)
        self.explorer_context_menu.add_command(label="Nouveau dossier", command=self.create_folder_from_explorer)
        self.explorer_context_menu.add_separator()
        self.explorer_context_menu.add_command(label="Supprimer", command=self.delete_selected_explorer_item)

    def create_resize_grips(self):
        grip_specs = [
            ("n", {"x": 0, "y": 0, "relwidth": 1, "height": self.resize_margin}, "size_ns"),
            ("s", {"x": 0, "rely": 1, "relwidth": 1, "height": self.resize_margin, "anchor": "sw"}, "size_ns"),
            ("w", {"x": 0, "y": 0, "width": self.resize_margin, "relheight": 1}, "size_we"),
            ("e", {"relx": 1, "y": 0, "width": self.resize_margin, "relheight": 1, "anchor": "ne"}, "size_we"),
            ("nw", {"x": 0, "y": 0, "width": self.resize_margin * 2, "height": self.resize_margin * 2}, "size_nw_se"),
            (
                "ne",
                {
                    "relx": 1,
                    "y": 0,
                    "width": self.resize_margin * 2,
                    "height": self.resize_margin * 2,
                    "anchor": "ne",
                },
                "size_ne_sw",
            ),
            (
                "sw",
                {
                    "x": 0,
                    "rely": 1,
                    "width": self.resize_margin * 2,
                    "height": self.resize_margin * 2,
                    "anchor": "sw",
                },
                "size_ne_sw",
            ),
            (
                "se",
                {
                    "relx": 1,
                    "rely": 1,
                    "width": self.resize_margin * 2,
                    "height": self.resize_margin * 2,
                    "anchor": "se",
                },
                "size_nw_se",
            ),
        ]

        for direction, place_options, cursor in grip_specs:
            grip = tk.Frame(self.root_frame, bd=0, highlightthickness=0, cursor=cursor)
            grip.place(**place_options)
            grip.bind("<Button-1>", lambda event, resize_direction=direction: self.start_window_resize(event, resize_direction))
            grip.bind("<B1-Motion>", self.resize_window)
            grip.bind("<ButtonRelease-1>", self.stop_window_resize)
            grip.lift()
            self.resize_grips.append(grip)

    def bind_shortcuts(self):
        self.bind("<Control-n>", lambda _event: self.add_new_poem())
        self.bind("<Control-o>", lambda _event: self.open_file())
        self.bind("<Control-s>", lambda _event: self.save_file())
        self.bind("<Control-k>", lambda _event: self.open_folder())

    def restore_custom_window_chrome(self, _event=None):
        if self.state() == "normal":
            self.after(10, lambda: self.overrideredirect(True))

    def start_window_drag(self, event):
        if self.is_maximized or self.get_resize_direction(event):
            return

        self.drag_offset_x = event.x_root - self.winfo_x()
        self.drag_offset_y = event.y_root - self.winfo_y()

    def drag_window(self, event):
        if self.is_maximized:
            return

        next_x = event.x_root - self.drag_offset_x
        next_y = event.y_root - self.drag_offset_y
        self.geometry(f"+{next_x}+{next_y}")

    def get_resize_direction(self, event) -> str:
        if self.is_maximized:
            return ""

        pointer_x = event.x_root - self.winfo_rootx()
        pointer_y = event.y_root - self.winfo_rooty()
        width = self.winfo_width()
        height = self.winfo_height()

        near_left = pointer_x <= self.resize_margin
        near_right = pointer_x >= width - self.resize_margin
        near_top = pointer_y <= self.resize_margin
        near_bottom = pointer_y >= height - self.resize_margin

        vertical = "n" if near_top else "s" if near_bottom else ""
        horizontal = "w" if near_left else "e" if near_right else ""

        return f"{vertical}{horizontal}"

    def start_window_resize(self, event, direction: str):
        if not direction or self.is_maximized:
            return

        self.resize_direction = direction
        self.resize_start_width = self.winfo_width()
        self.resize_start_height = self.winfo_height()
        self.resize_start_window_x = self.winfo_x()
        self.resize_start_window_y = self.winfo_y()
        self.resize_start_x = event.x_root
        self.resize_start_y = event.y_root

    def resize_window(self, event):
        if not self.resize_direction:
            return

        min_width, min_height = self.minsize()
        delta_x = event.x_root - self.resize_start_x
        delta_y = event.y_root - self.resize_start_y
        next_x = self.resize_start_window_x
        next_y = self.resize_start_window_y
        next_width = self.resize_start_width
        next_height = self.resize_start_height

        if "e" in self.resize_direction:
            next_width = max(min_width, self.resize_start_width + delta_x)

        if "s" in self.resize_direction:
            next_height = max(min_height, self.resize_start_height + delta_y)

        if "w" in self.resize_direction:
            requested_width = self.resize_start_width - delta_x
            next_width = max(min_width, requested_width)
            next_x = self.resize_start_window_x + self.resize_start_width - next_width

        if "n" in self.resize_direction:
            requested_height = self.resize_start_height - delta_y
            next_height = max(min_height, requested_height)
            next_y = self.resize_start_window_y + self.resize_start_height - next_height

        self.geometry(f"{next_width}x{next_height}+{next_x}+{next_y}")

    def stop_window_resize(self, _event=None):
        self.resize_direction = ""

    def minimize_window(self):
        self.overrideredirect(False)
        self.iconify()

    def toggle_maximize_window(self):
        if self.is_maximized:
            if self.restore_geometry:
                self.geometry(self.restore_geometry)
            self.is_maximized = False
            return

        self.restore_geometry = self.geometry()
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
        self.is_maximized = True

    def toggle_theme(self):
        self.dark_theme_enabled.set(not self.dark_theme_enabled.get())
        self.apply_theme()
        self.save_app_settings()

    def get_settings_path(self) -> str:
        app_data_path = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        settings_folder = os.path.join(app_data_path, "PoetryEditor")
        return os.path.join(settings_folder, self.SETTINGS_FILENAME)

    def load_app_settings(self) -> dict:
        settings_path = self.get_settings_path()

        if not os.path.exists(settings_path):
            return {}

        try:
            with open(settings_path, "r", encoding="utf-8") as settings_file:
                settings = json.load(settings_file)
        except (OSError, json.JSONDecodeError):
            return {}

        return settings if isinstance(settings, dict) else {}

    def get_saved_metric_objective(self) -> str:
        metric_objective = self.app_settings.get("metric_objective", "Libre")
        return metric_objective if metric_objective in Editor.get_metric_names() else "Libre"

    def save_app_settings(self):
        settings_path = self.get_settings_path()
        os.makedirs(os.path.dirname(settings_path), exist_ok=True)

        settings = {
            "dark_theme_enabled": self.dark_theme_enabled.get(),
            "current_folder": self.current_folder,
            "current_file": self.editor_core.file_path,
            "metric_objective": self.metric_objective.get(),
        }

        try:
            with open(settings_path, "w", encoding="utf-8") as settings_file:
                json.dump(settings, settings_file, indent=2, ensure_ascii=False)
        except OSError:
            pass

    def restore_session(self):
        folder_path = self.app_settings.get("current_folder")
        file_path = self.app_settings.get("current_file")

        if folder_path and os.path.isdir(folder_path):
            self.current_folder = folder_path
            self.folder_label.configure(text=folder_path)
            self.populate_folder_tree(folder_path)

        if file_path and os.path.isfile(file_path):
            self.load_file(file_path, persist_settings=False)

    def new_file(self):
        if self.confirm_unsaved_changes():
            self.text_edit.delete("1.0", tk.END)
            self.text_edit.edit_modified(False)
            self.editor_core = Editor()
            self.clear_current_image()
            self.clear_poem_structure()
            self.clear_syllable_counts()
            self.update_window_title()
            self.update_status()
            self.save_app_settings()

    def add_new_poem(self):
        if not self.confirm_unsaved_changes():
            return

        dialog = self.open_poem_creation_dialog()

        if not dialog.result:
            return

        content = dialog.result["content"]
        poem_type = dialog.result["poem_type"]
        self.text_edit.delete("1.0", tk.END)
        self.text_edit.insert("1.0", content)
        self.text_edit.edit_modified(False)
        self.editor_core = Editor()
        self.editor_core.set_content(content)
        self.editor_core.mark_modified()
        self.metric_objective.set(poem_type.get("metric", "Libre"))
        self.clear_current_image()
        self.display_poem_structure(poem_type, content)
        self.clear_syllable_counts()
        self.update_window_title()
        self.update_status()
        self.focus_editor_after_poem_creation()
        self.save_app_settings()

    def open_poem_creation_dialog(self) -> PoemCreationDialog:
        theme_name = "dark" if self.dark_theme_enabled.get() else "light"
        return PoemCreationDialog(self, POEM_TYPES, self.THEMES[theme_name])

    def focus_editor_after_poem_creation(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self.text_edit.focus_set()
        self.text_edit.mark_set(tk.INSERT, "end-1c")
        self.text_edit.see(tk.INSERT)

    def display_poem_structure(self, poem_type: dict, content: str):
        self.current_poem_type = poem_type
        self.structure_type_label.configure(text=poem_type["name"])
        self.structure_metric_label.configure(text=f"Objectif : {poem_type.get('metric', 'Libre')}")
        self.structure_rules_text.configure(text=poem_type["rules"])
        self.refresh_poem_structure_text(content)

        if not self.structure_panel.winfo_ismapped():
            self.structure_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(18, 0))

    def refresh_poem_structure_text(self, content: str | None = None):
        if not self.current_poem_type:
            return

        content = self.get_text_content() if content is None else content
        structure_text = build_editor_structure_text(self.current_poem_type, content)
        self.structure_text.configure(state=tk.NORMAL)
        self.structure_text.delete("1.0", tk.END)
        self.structure_text.insert("1.0", structure_text)
        self.structure_text.configure(state=tk.DISABLED)

    def update_poem_structure_from_content(self, content: str):
        poem_type = detect_poem_type_from_content(content)

        if poem_type:
            self.display_poem_structure(poem_type, content)
            return

        self.clear_poem_structure()

    def hide_poem_structure(self):
        self.structure_panel.pack_forget()

    def clear_poem_structure(self):
        self.current_poem_type = None
        self.structure_text.configure(state=tk.NORMAL)
        self.structure_text.delete("1.0", tk.END)
        self.structure_text.configure(state=tk.DISABLED)
        self.structure_panel.pack_forget()

    def open_file(self):
        if not self.confirm_unsaved_changes():
            return

        path = filedialog.askopenfilename(
            title="Ouvrir un fichier",
            filetypes=self.TEXT_FILETYPES,
        )

        if path:
            self.load_file(path)

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
            self.refresh_folder_tree(self.editor_core.file_path)
            self.save_app_settings()

        return success

    def save_file_as(self):
        path = filedialog.asksaveasfilename(
            title="Sauvegarder sous",
            defaultextension=".txt",
            filetypes=self.TEXT_FILETYPES,
        )

        if not path:
            return False

        content = self.get_text_content()
        success = self.file_service.write(path, content)

        if success:
            self.editor_core.set_file_path(path)
            self.editor_core.set_content(content)
            self.text_edit.edit_modified(False)
            self.load_associated_image()
            self.update_window_title()
            self.update_status()
            self.refresh_folder_tree(path)
            self.save_app_settings()

        return success

    def open_folder(self):
        path = filedialog.askdirectory(title="Ouvrir un dossier")

        if not path:
            return

        self.current_folder = path
        self.folder_label.configure(text=path)
        self.populate_folder_tree(path)
        self.save_app_settings()

    def show_explorer_context_menu(self, event):
        item_id = self.folder_tree.identify_row(event.y)

        if item_id:
            self.folder_tree.selection_set(item_id)
            self.folder_tree.focus(item_id)

        self.explorer_context_menu.tk_popup(event.x_root, event.y_root)

    def create_text_from_explorer(self):
        target_folder = self.get_explorer_target_folder()

        if not target_folder:
            messagebox.showinfo("Explorateur", "Ouvrez un dossier avant de creer un texte.", parent=self)
            return

        name = simpledialog.askstring("Nouveau texte", "Nom du texte :", parent=self)

        if not name:
            return

        filename = name.strip()

        if not filename:
            return

        if not os.path.splitext(filename)[1]:
            filename = f"{filename}.txt"

        path = os.path.join(target_folder, filename)

        if os.path.exists(path):
            messagebox.showerror("Nouveau texte", "Un fichier avec ce nom existe deja.", parent=self)
            return

        try:
            with open(path, "w", encoding="utf-8"):
                pass
        except OSError as error:
            messagebox.showerror("Nouveau texte", f"Impossible de creer le texte: {error}", parent=self)
            return

        self.refresh_folder_tree(path)
        self.load_file(path)

    def create_poem_from_explorer(self):
        target_folder = self.get_explorer_target_folder()

        if not target_folder:
            messagebox.showinfo("Explorateur", "Ouvrez un dossier avant de creer un poeme.", parent=self)
            return

        dialog = self.open_poem_creation_dialog()

        if not dialog.result:
            return

        filename = f"{slugify_filename(dialog.result['title'])}.txt"
        path = os.path.join(target_folder, filename)
        suffix = 2

        while os.path.exists(path):
            path = os.path.join(target_folder, f"{slugify_filename(dialog.result['title'])}_{suffix}.txt")
            suffix += 1

        success = self.file_service.write(path, dialog.result["content"])

        if not success:
            messagebox.showerror("Ajouter un poeme", "Impossible de creer le poeme.", parent=self)
            return

        self.refresh_folder_tree(path)
        self.load_file(path)
        self.metric_objective.set(dialog.result["poem_type"].get("metric", "Libre"))
        self.focus_editor_after_poem_creation()
        self.save_app_settings()

    def create_folder_from_explorer(self):
        target_folder = self.get_explorer_target_folder()

        if not target_folder:
            messagebox.showinfo("Explorateur", "Ouvrez un dossier avant de creer un dossier.", parent=self)
            return

        name = simpledialog.askstring("Nouveau dossier", "Nom du dossier :", parent=self)

        if not name:
            return

        folder_name = name.strip()

        if not folder_name:
            return

        path = os.path.join(target_folder, folder_name)

        if os.path.exists(path):
            messagebox.showerror("Nouveau dossier", "Un element avec ce nom existe deja.", parent=self)
            return

        try:
            os.makedirs(path)
        except OSError as error:
            messagebox.showerror("Nouveau dossier", f"Impossible de creer le dossier: {error}", parent=self)
            return

        self.refresh_folder_tree(path)

    def delete_selected_explorer_item(self):
        path = self.get_selected_explorer_path()

        if not path:
            messagebox.showinfo("Explorateur", "Selectionnez un texte ou un dossier a supprimer.", parent=self)
            return

        if self.current_folder and self.paths_match(path, self.current_folder):
            messagebox.showinfo("Explorateur", "Le dossier ouvert ne peut pas etre supprime ici.", parent=self)
            return

        if self.is_current_file_affected_by_delete(path):
            if not self.confirm_unsaved_changes():
                return

        item_name = os.path.basename(path)
        is_folder = os.path.isdir(path)
        message = f"Supprimer definitivement le dossier '{item_name}' et son contenu ?"

        if not is_folder:
            message = f"Supprimer definitivement le texte '{item_name}' ?"

        if not messagebox.askyesno("Supprimer", message, parent=self):
            return

        try:
            if is_folder:
                shutil.rmtree(path)
            else:
                self.delete_associated_image(path)
                os.remove(path)
        except OSError as error:
            messagebox.showerror("Supprimer", f"Impossible de supprimer: {error}", parent=self)
            return

        if self.editor_core.file_path and not os.path.exists(self.editor_core.file_path):
            self.text_edit.delete("1.0", tk.END)
            self.text_edit.edit_modified(False)
            self.editor_core = Editor()
            self.clear_poem_structure()
            self.clear_syllable_counts()
            self.update_window_title()
            self.update_status()

        self.refresh_folder_tree()
        self.save_app_settings()

    def import_image_for_current_file(self):
        if self.image_generation_pending:
            return

        poem = self.get_text_content().strip()

        if not poem:
            messagebox.showinfo("Image IA", "Ecrivez un poeme avant de generer une image.", parent=self)
            return

        if not self.editor_core.has_file():
            if not self.save_file_as():
                return

        text_path = self.editor_core.file_path
        prompt = self.build_image_generation_prompt(poem)
        destination_path = self.get_generated_image_path(text_path)
        self.image_generation_id += 1
        generation_id = self.image_generation_id

        self.set_image_generation_state(True)
        self.after(45000, lambda: self.cancel_stalled_image_generation(generation_id))

        thread = threading.Thread(
            target=self.generate_image_for_file,
            args=(generation_id, text_path, prompt, destination_path),
            daemon=True,
        )
        thread.start()

    def generate_image_for_file(self, generation_id: int, text_path: str, prompt: str, destination_path: str):
        error_message = ""

        try:
            self.pollinations_service.generate_image(prompt, destination_path)

            metadata = self.read_image_metadata(text_path)
            metadata[os.path.basename(text_path)] = os.path.relpath(destination_path, os.path.dirname(text_path))
            self.write_image_metadata(text_path, metadata)
        except Exception as error:
            error_message = str(error)

        self.after(
            0,
            lambda: self.finish_image_generation(generation_id, text_path, destination_path, error_message),
        )

    def finish_image_generation(
        self,
        generation_id: int,
        text_path: str,
        destination_path: str,
        error_message: str,
    ):
        if generation_id != self.image_generation_id:
            return

        self.set_image_generation_state(False)

        if error_message:
            messagebox.showerror("Image IA", f"Impossible de generer l'image: {error_message}", parent=self)
            return

        if self.editor_core.file_path != text_path:
            return

        self.current_image_path = destination_path
        self.display_current_image()

    def cancel_stalled_image_generation(self, generation_id: int):
        if not self.image_generation_pending or generation_id != self.image_generation_id:
            return

        self.image_generation_id += 1
        self.set_image_generation_state(False)
        messagebox.showerror(
            "Image IA",
            "La generation prend trop de temps. Reessayez dans quelques instants.",
            parent=self,
        )

    def set_image_generation_state(self, is_pending: bool):
        self.image_generation_pending = is_pending

        if self.generate_image_button is not None:
            self.generate_image_button.configure(
                text="Generation...",
                state=tk.DISABLED,
            )

            if not is_pending:
                self.generate_image_button.configure(
                    text="Image IA",
                    state=tk.NORMAL,
                )

        if is_pending:
            self.status_left.configure(text="Generation de l'image IA en cours...")
        else:
            self.update_status()

    def build_image_generation_prompt(self, poem: str) -> str:
        poem_excerpt = " ".join(poem.split())[:1200]
        return (
            "Poetic book illustration inspired by this poem, expressive, atmospheric, painterly, "
            "high detail, harmonious composition, no visible text. Poem: "
            f"{poem_excerpt}"
        )

    def remove_image_from_current_file(self):
        if not self.editor_core.has_file():
            return

        self.delete_associated_image(self.editor_core.file_path)
        self.clear_current_image()

    def delete_associated_image(self, text_path: str):
        metadata = self.read_image_metadata(text_path)
        text_key = os.path.basename(text_path)

        if text_key not in metadata:
            return

        image_path = self.get_associated_image_path(text_path, metadata)
        metadata.pop(text_key, None)
        self.write_image_metadata(text_path, metadata)

        if image_path and os.path.exists(image_path):
            os.remove(image_path)

    def get_generated_image_path(self, text_path: str) -> str:
        text_folder = os.path.dirname(text_path)
        text_name = os.path.splitext(os.path.basename(text_path))[0]
        timestamp = int(time.time() * 1000)
        assets_folder = os.path.join(text_folder, self.IMAGE_ASSETS_FOLDER)
        return os.path.join(assets_folder, f"{text_name}_generated_{timestamp}.jpg")

    def read_image_metadata(self, text_path: str) -> dict[str, str]:
        metadata_path = self.get_image_metadata_path(text_path)

        if not os.path.exists(metadata_path):
            return {}

        try:
            with open(metadata_path, "r", encoding="utf-8") as metadata_file:
                metadata = json.load(metadata_file)
        except (OSError, json.JSONDecodeError):
            return {}

        return metadata if isinstance(metadata, dict) else {}

    def write_image_metadata(self, text_path: str, metadata: dict[str, str]):
        metadata_path = self.get_image_metadata_path(text_path)

        with open(metadata_path, "w", encoding="utf-8") as metadata_file:
            json.dump(metadata, metadata_file, indent=2, ensure_ascii=False)

    def get_image_metadata_path(self, text_path: str) -> str:
        return os.path.join(os.path.dirname(text_path), self.IMAGE_METADATA_FILENAME)

    def get_associated_image_path(self, text_path: str, metadata: dict[str, str] | None = None) -> str:
        metadata = metadata if metadata is not None else self.read_image_metadata(text_path)
        relative_image_path = metadata.get(os.path.basename(text_path), "")

        if not relative_image_path:
            return ""

        image_path = os.path.join(os.path.dirname(text_path), relative_image_path)
        return image_path if os.path.exists(image_path) else ""

    def load_associated_image(self):
        if not self.editor_core.has_file():
            self.clear_current_image()
            return

        self.current_image_path = self.get_associated_image_path(self.editor_core.file_path)
        self.display_current_image()

    def display_current_image(self):
        if not self.current_image_path:
            self.clear_current_image()
            return

        try:
            image = Image.open(self.current_image_path)
            image.thumbnail((320, 460), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(image)
        except (OSError, tk.TclError):
            self.clear_current_image()
            return

        self.image_preview_label.configure(image=self.image_preview)
        self.image_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(18, 0))

    def clear_current_image(self):
        self.current_image_path = None
        self.image_preview = None
        self.image_preview_label.configure(image="")
        self.image_panel.pack_forget()

    def get_selected_explorer_path(self) -> str:
        selection = self.folder_tree.selection()

        if not selection:
            return ""

        return self.get_tree_item_path(selection[0])

    def get_explorer_target_folder(self) -> str:
        selected_path = self.get_selected_explorer_path()

        if selected_path:
            if os.path.isdir(selected_path):
                return selected_path

            return os.path.dirname(selected_path)

        return self.current_folder or ""

    def is_current_file_affected_by_delete(self, deleted_path: str) -> bool:
        if not self.editor_core.file_path:
            return False

        if self.paths_match(deleted_path, self.editor_core.file_path):
            return True

        if os.path.isdir(deleted_path):
            return self.is_path_inside_folder(self.editor_core.file_path, deleted_path)

        return False

    def paths_match(self, first_path: str, second_path: str) -> bool:
        return os.path.normcase(os.path.abspath(first_path)) == os.path.normcase(os.path.abspath(second_path))

    def is_path_inside_folder(self, path: str, folder: str) -> bool:
        try:
            common_path = os.path.commonpath(
                [
                    os.path.abspath(path),
                    os.path.abspath(folder),
                ]
            )
        except ValueError:
            return False

        return os.path.normcase(common_path) == os.path.normcase(os.path.abspath(folder))

    def populate_folder_tree(self, path: str):
        self.folder_tree.delete(*self.folder_tree.get_children())
        root_id = self.folder_tree.insert(
            "",
            tk.END,
            text=os.path.basename(path) or path,
            values=(path,),
            open=True,
        )
        self.insert_folder_children(root_id, path)

    def refresh_folder_tree(self, selected_path: str | None = None):
        if not self.current_folder:
            return

        selected_path = selected_path or self.editor_core.file_path

        if selected_path and not self.is_path_in_current_folder(selected_path):
            return

        open_paths = self.get_open_tree_paths()
        self.populate_folder_tree(self.current_folder)

        for path in open_paths:
            item_id = self.find_tree_item_by_path(path)

            if item_id:
                self.ensure_tree_folder_loaded(item_id, path)
                self.folder_tree.item(item_id, open=True)

        if selected_path:
            selected_item = self.find_tree_item_by_path(selected_path)

            if selected_item:
                self.folder_tree.selection_set(selected_item)
                self.folder_tree.focus(selected_item)
                self.folder_tree.see(selected_item)

    def insert_folder_children(self, parent_id: str, path: str):
        try:
            entries = sorted(
                os.scandir(path),
                key=lambda entry: (not entry.is_dir(), entry.name.lower()),
            )
        except OSError:
            return

        for entry in entries:
            if entry.name.startswith("."):
                continue

            label = self.get_tree_display_name(entry)
            item_id = self.folder_tree.insert(
                parent_id,
                tk.END,
                text=label,
                values=(entry.path,),
            )

            if entry.is_dir():
                self.folder_tree.insert(item_id, tk.END, text="Chargement...", values=("",))

    def get_tree_display_name(self, entry: os.DirEntry) -> str:
        if entry.is_dir():
            return entry.name

        name_without_extension, _extension = os.path.splitext(entry.name)
        return name_without_extension or entry.name

    def on_tree_open(self, _event=None):
        item_id = self.folder_tree.focus()
        path = self.get_tree_item_path(item_id)

        if not path or not os.path.isdir(path):
            return

        children = self.folder_tree.get_children(item_id)

        if len(children) == 1 and not self.get_tree_item_path(children[0]):
            self.folder_tree.delete(children[0])
            self.insert_folder_children(item_id, path)

    def open_selected_tree_file(self, _event=None):
        item_id = self.folder_tree.focus()
        path = self.get_tree_item_path(item_id)

        if not path or os.path.isdir(path):
            return

        if self.confirm_unsaved_changes():
            self.load_file(path)

    def get_tree_item_path(self, item_id: str) -> str:
        if not item_id:
            return ""

        values = self.folder_tree.item(item_id, "values")
        return values[0] if values else ""

    def get_open_tree_paths(self) -> set[str]:
        open_paths = set()

        def collect_open_paths(parent_id: str):
            for item_id in self.folder_tree.get_children(parent_id):
                path = self.get_tree_item_path(item_id)

                if path and os.path.isdir(path) and self.folder_tree.item(item_id, "open"):
                    open_paths.add(path)
                    collect_open_paths(item_id)

        collect_open_paths("")
        return open_paths

    def ensure_tree_folder_loaded(self, item_id: str, path: str):
        children = self.folder_tree.get_children(item_id)

        if len(children) == 1 and not self.get_tree_item_path(children[0]):
            self.folder_tree.delete(children[0])
            self.insert_folder_children(item_id, path)

    def find_tree_item_by_path(self, target_path: str) -> str:
        root_items = self.folder_tree.get_children("")

        if not root_items or not self.current_folder:
            return ""

        root_id = root_items[0]
        root_path = os.path.abspath(self.current_folder)
        target_path = os.path.abspath(target_path)

        if os.path.normcase(root_path) == os.path.normcase(target_path):
            return root_id

        try:
            relative_path = os.path.relpath(target_path, root_path)
        except ValueError:
            return ""

        if relative_path.startswith(".."):
            return ""

        current_id = root_id
        current_path = root_path

        for part in relative_path.split(os.sep):
            self.ensure_tree_folder_loaded(current_id, current_path)
            next_id = ""

            for child_id in self.folder_tree.get_children(current_id):
                child_path = self.get_tree_item_path(child_id)

                if child_path and os.path.basename(child_path) == part:
                    next_id = child_id
                    current_path = child_path
                    break

            if not next_id:
                return ""

            current_id = next_id

        return current_id

    def is_path_in_current_folder(self, path: str) -> bool:
        if not self.current_folder:
            return False

        return self.is_path_inside_folder(path, self.current_folder)

    def load_file(self, path: str, persist_settings: bool = True):
        content = self.file_service.read(path)
        self.editor_core.set_content(content)
        self.editor_core.set_file_path(path)

        self.text_edit.delete("1.0", tk.END)
        self.text_edit.insert("1.0", content)
        self.text_edit.edit_modified(False)
        self.load_associated_image()
        self.update_poem_structure_from_content(content)
        self.clear_syllable_counts()
        self.update_window_title()
        self.update_status()

        if persist_settings:
            self.save_app_settings()

    def on_text_changed(self, _event=None):
        if self.text_edit.edit_modified():
            self.editor_core.mark_modified()
            self.clear_syllable_counts()
            self.refresh_poem_structure_text()
            self.update_window_title()
            self.update_status()
            self.text_edit.edit_modified(False)

    def apply_theme(self):
        theme_name = "dark" if self.dark_theme_enabled.get() else "light"
        theme = self.THEMES[theme_name]

        self.configure(bg=theme["window_bg"])
        self.root_frame.configure(bg=theme["window_bg"])
        self.window_chrome.configure(bg=theme["chrome_bg"], highlightthickness=1, highlightbackground=theme["chrome_border"])
        self.window_title_area.configure(bg=theme["chrome_bg"])
        self.app_mark.configure(bg=theme["chrome_mark_bg"], fg=theme["chrome_mark_fg"])
        self.window_text_block.configure(bg=theme["chrome_bg"])
        self.window_app_name.configure(bg=theme["chrome_bg"], fg=theme["editor_fg"])
        self.window_actions.configure(bg=theme["chrome_bg"])
        for grip in self.resize_grips:
            grip.configure(bg=theme["window_bg"])
        self.toolbar.configure(bg=theme["window_bg"])
        self.title_block.configure(bg=theme["window_bg"])
        self.app_title.configure(bg=theme["window_bg"], fg=theme["editor_fg"])
        self.file_label.configure(bg=theme["window_bg"], fg=theme["muted_fg"])
        self.workspace.configure(bg=theme["window_bg"])

        self.sidebar_shell.configure(
            bg=theme["surface_bg"],
            highlightbackground=theme["surface_border"],
            highlightcolor=theme["surface_border"],
        )
        self.sidebar_header.configure(bg=theme["surface_bg"])
        self.sidebar_title.configure(bg=theme["surface_bg"], fg=theme["editor_fg"])
        self.folder_label.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.tree_frame.configure(bg=theme["surface_bg"])

        self.editor_shell.configure(
            bg=theme["surface_bg"],
            highlightbackground=theme["surface_border"],
            highlightcolor=theme["surface_border"],
        )
        self.editor_header.configure(bg=theme["surface_bg"])
        self.editor_title_block.configure(bg=theme["surface_bg"])
        self.mode_label.configure(bg=theme["surface_bg"], fg=theme["editor_fg"])
        self.analysis_bar.configure(bg=theme["surface_bg"])
        self.analysis_label.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.metric_frame.configure(bg=theme["surface_bg"])
        self.metric_label.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.editor_tools.configure(bg=theme["surface_bg"])
        self.hint_label.configure(bg=theme["surface_bg"], fg=theme["muted_fg"])
        self.editor_content.configure(bg=theme["surface_bg"])
        self.editor_body.configure(bg=theme["surface_bg"])
        self.structure_panel.configure(
            bg=theme["toolbar_bg"],
            highlightbackground=theme["surface_border"],
            highlightcolor=theme["surface_border"],
        )
        self.structure_header.configure(bg=theme["toolbar_bg"])
        self.structure_title.configure(bg=theme["toolbar_bg"], fg=theme["editor_fg"])
        self.structure_type_label.configure(bg=theme["toolbar_bg"], fg=theme["editor_fg"])
        self.structure_metric_label.configure(bg=theme["toolbar_bg"], fg=theme["muted_fg"])
        self.structure_rules_text.configure(bg=theme["toolbar_bg"], fg=theme["muted_fg"])
        self.structure_list_label.configure(bg=theme["toolbar_bg"], fg=theme["editor_fg"])
        self.structure_text_frame.configure(bg=theme["editor_bg"])
        self.structure_text.configure(
            bg=theme["editor_bg"],
            fg=theme["editor_fg"],
            insertbackground=theme["insert_bg"],
            selectbackground=theme["select_bg"],
            selectforeground=theme["select_fg"],
        )
        self.image_panel.configure(bg=theme["surface_bg"])
        self.image_preview_label.configure(bg=theme["surface_bg"])
        self.status_bar.configure(bg=theme["window_bg"])
        self.status_left.configure(bg=theme["window_bg"], fg=theme["muted_fg"])
        self.status_right.configure(bg=theme["window_bg"], fg=theme["muted_fg"])

        self.text_edit.configure(
            bg=theme["editor_bg"],
            fg=theme["editor_fg"],
            insertbackground=theme["insert_bg"],
            selectbackground=theme["select_bg"],
            selectforeground=theme["select_fg"],
        )
        self.syllable_gutter.configure(bg=theme["editor_bg"])
        self.ui_style.configure(
            self.folder_tree_style_name,
            background=theme["surface_bg"],
            foreground=theme["editor_fg"],
            fieldbackground=theme["surface_bg"],
            borderwidth=0,
            rowheight=24,
            font=("Segoe UI", 9),
            relief="flat",
        )
        self.ui_style.map(
            self.folder_tree_style_name,
            background=[("selected", theme["tree_selected_bg"])],
            foreground=[("selected", theme["editor_fg"])],
        )
        self.ui_style.configure(
            self.metric_selector_style_name,
            fieldbackground=theme["button_bg"],
            background=theme["button_bg"],
            foreground=theme["button_fg"],
            arrowcolor=theme["muted_fg"],
            bordercolor=theme["surface_border"],
            darkcolor=theme["button_bg"],
            lightcolor=theme["button_bg"],
            selectbackground=theme["button_active_bg"],
            selectforeground=theme["button_fg"],
        )
        self.ui_style.map(
            self.metric_selector_style_name,
            fieldbackground=[
                ("readonly", theme["button_bg"]),
                ("disabled", theme["button_bg"]),
            ],
            background=[
                ("readonly", theme["button_bg"]),
                ("active", theme["button_active_bg"]),
            ],
            foreground=[
                ("readonly", theme["button_fg"]),
                ("disabled", theme["muted_fg"]),
            ],
            arrowcolor=[
                ("readonly", theme["muted_fg"]),
                ("active", theme["editor_fg"]),
            ],
        )
        self.ui_style.configure(
            self.scrollbar_style_name,
            background=theme["scrollbar_bg"],
            darkcolor=theme["scrollbar_bg"],
            lightcolor=theme["scrollbar_bg"],
            troughcolor=theme["surface_bg"],
            bordercolor=theme["surface_bg"],
            arrowcolor=theme["muted_fg"],
            gripcount=0,
            width=10,
            relief="flat",
        )
        self.ui_style.map(
            self.scrollbar_style_name,
            background=[("active", theme["scrollbar_active_bg"])],
            arrowcolor=[("active", theme["editor_fg"])],
        )

        for button in self.toolbar_buttons:
            button.configure(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["button_active_bg"],
                activeforeground=theme["button_fg"],
            )

        self.update_analysis_button_styles(theme)

        for index, button in enumerate(self.window_control_buttons):
            active_bg = theme["chrome_close_hover"] if index == 2 else theme["chrome_button_hover"]
            active_fg = "#ffffff" if index == 2 else theme["editor_fg"]
            button.configure(
                bg=theme["chrome_bg"],
                fg=theme["editor_fg"],
                activebackground=active_bg,
                activeforeground=active_fg,
            )

        for menu in self.menus:
            menu.configure(
                bg=theme["menu_bg"],
                fg=theme["menu_fg"],
                activebackground=theme["active_bg"],
                activeforeground=theme["active_fg"],
                selectcolor=theme["editor_bg"],
            )

        self.redraw_syllable_gutter()

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

    def show_syllable_count(self):
        content = self.get_text_content()
        self.analysis_mode.set("syllables")
        self.update_analysis_button_styles()
        self.syllable_count_pending = True
        self.syllable_line_counts = []
        self.syllable_line_targets = []
        self.rhyme_line_labels = []
        self.redraw_syllable_gutter()

        thread = threading.Thread(
            target=self.calculate_syllable_count,
            args=(content,),
            daemon=True,
        )
        thread.start()

    def update_analysis_button_styles(self, theme: dict | None = None):
        if theme is None:
            theme_name = "dark" if self.dark_theme_enabled.get() else "light"
            theme = self.THEMES[theme_name]

        active_mode = self.analysis_mode.get()
        analysis_buttons = (
            (self.syllables_button, "syllables"),
            (self.rhymes_button, "rhymes"),
        )

        for button, mode in analysis_buttons:
            is_active = active_mode == mode
            button.configure(
                bg=theme["button_active_bg"] if is_active else theme["button_bg"],
                fg=theme["active_fg"] if is_active else theme["button_fg"],
                activebackground=theme["button_active_bg"],
                activeforeground=theme["active_fg"],
            )

    def calculate_syllable_count(self, content: str):
        try:
            line_counts = self.editor_core.count_line_syllables(content)
            total = sum(line_counts)
        except Exception as error:
            error_message = str(error)
            self.after(0, lambda: self.show_syllable_error(error_message))
            return

        self.after(0, lambda: self.display_syllable_count(content, total, line_counts))

    def display_syllable_count(self, content: str, total: int, line_counts: list[int]):
        self.syllable_count_pending = False
        self.syllable_line_counts = line_counts
        self.syllable_line_targets = self.editor_core.get_line_metric_targets(
            self.metric_objective.get(),
            line_counts,
        )
        self.update_syllable_status(total)
        self.redraw_syllable_gutter()

    def show_syllable_error(self, error_message: str):
        self.syllable_count_pending = False
        self.syllable_line_counts = []
        self.syllable_line_targets = []
        self.redraw_syllable_gutter()
        messagebox.showerror("Syllabes", f"Impossible de compter les syllabes: {error_message}", parent=self)

    def clear_syllable_counts(self):
        self.analysis_mode.set("")
        self.syllable_count_pending = False
        self.syllable_line_counts = []
        self.syllable_line_targets = []
        self.rhyme_line_labels = []
        self.update_analysis_button_styles()
        self.redraw_syllable_gutter()

    def show_rhyme_scheme(self):
        content = self.get_text_content()
        self.analysis_mode.set("rhymes")
        self.update_analysis_button_styles()
        self.syllable_count_pending = False
        self.syllable_line_counts = []
        self.syllable_line_targets = []
        self.rhyme_line_labels = self.editor_core.get_line_rhyme_labels(content)
        scheme = self.editor_core.get_rhyme_scheme(self.rhyme_line_labels)

        if scheme:
            self.status_left.configure(text=f"Schema de rimes: {scheme}")
        else:
            self.status_left.configure(text="Aucune rime detectee")

        self.redraw_syllable_gutter()

    def on_metric_objective_changed(self, _event=None):
        if self.syllable_line_counts:
            total = sum(self.syllable_line_counts)
            self.syllable_line_targets = self.editor_core.get_line_metric_targets(
                self.metric_objective.get(),
                self.syllable_line_counts,
            )
            self.update_syllable_status(total)
            self.redraw_syllable_gutter()

        self.save_app_settings()

    def update_syllable_status(self, total: int):
        metric_name = self.metric_objective.get()
        status = f"Total: {total} syllabe{'s' if total > 1 else ''}"

        if metric_name != "Libre":
            matching_lines, checked_lines = self.editor_core.summarize_metric_progress(
                metric_name,
                self.syllable_line_counts,
            )
            status = f"{status}   Objectif {metric_name}: {matching_lines}/{checked_lines} lignes"

        self.status_left.configure(text=status)

    def on_editor_navigation(self, _event=None):
        self.update_status()
        self.schedule_syllable_gutter_redraw()

    def on_text_scroll(self, first, last):
        self.scrollbar.set(first, last)
        self.schedule_syllable_gutter_redraw()

    def scroll_text(self, *args):
        self.text_edit.yview(*args)
        self.schedule_syllable_gutter_redraw()

    def schedule_syllable_gutter_redraw(self):
        self.after_idle(self.redraw_syllable_gutter)

    def redraw_syllable_gutter(self):
        self.syllable_gutter.delete("all")

        theme_name = "dark" if self.dark_theme_enabled.get() else "light"
        theme = self.THEMES[theme_name]
        fg = theme["muted_fg"]

        if self.syllable_count_pending:
            line_index = self.text_edit.index("@0,0").split(".")[0]
            info = self.text_edit.dlineinfo(f"{line_index}.0")

            if info:
                _x, y, _width, height, _baseline = info
                self.syllable_gutter.create_text(
                    19,
                    y + height // 2,
                    text="...",
                    fill=fg,
                    font=("Segoe UI", 9, "bold"),
                )

            return

        if self.rhyme_line_labels:
            self.redraw_rhyme_gutter(theme)
            return

        if not self.syllable_line_counts:
            return

        index = self.text_edit.index("@0,0")

        while True:
            line = int(index.split(".")[0])
            info = self.text_edit.dlineinfo(f"{line}.0")

            if info is None:
                break

            if line <= len(self.syllable_line_counts):
                _x, y, _width, height, _baseline = info
                count = self.syllable_line_counts[line - 1]
                text = str(count) if count else ""
                target = self.syllable_line_targets[line - 1] if line <= len(self.syllable_line_targets) else None
                line_fg = fg

                if target is not None and count:
                    line_fg = theme["metric_ok_fg"] if count == target else theme["metric_error_fg"]

                if text:
                    self.syllable_gutter.create_text(
                        19,
                        y + height // 2,
                        text=text,
                        fill=line_fg,
                        font=("Segoe UI", 9, "bold"),
                    )

            next_index = self.text_edit.index(f"{line + 1}.0")

            if next_index == index:
                break

            index = next_index

    def redraw_rhyme_gutter(self, theme: dict):
        index = self.text_edit.index("@0,0")
        palette = theme["rhyme_palette"]

        while True:
            line = int(index.split(".")[0])
            info = self.text_edit.dlineinfo(f"{line}.0")

            if info is None:
                break

            if line <= len(self.rhyme_line_labels):
                _x, y, _width, height, _baseline = info
                label = self.rhyme_line_labels[line - 1]

                if label:
                    color_index = self.get_rhyme_color_index(label)
                    self.syllable_gutter.create_text(
                        19,
                        y + height // 2,
                        text=label,
                        fill=palette[color_index % len(palette)],
                        font=("Segoe UI", 9, "bold"),
                    )

            next_index = self.text_edit.index(f"{line + 1}.0")

            if next_index == index:
                break

            index = next_index

    def get_rhyme_color_index(self, label: str) -> int:
        color_index = 0

        for char in label:
            color_index = color_index * 26 + (ord(char) - ord("A") + 1)

        return max(0, color_index - 1)

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
            self.save_app_settings()
            self.destroy()

    def get_text_content(self) -> str:
        return self.text_edit.get("1.0", "end-1c")
