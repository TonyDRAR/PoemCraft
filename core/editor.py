class Editor:
    def __init__(self):
        self.content = ""
        self.file_path = None
        self.modified = False

    def set_content(self, text: str):
        self.content = text
        self.modified = False

    def get_content(self) -> str:
        return self.content

    def set_file_path(self, path: str):
        self.file_path = path

    def mark_modified(self):
        self.modified = True

    def is_modified(self) -> bool:
        return self.modified

    def has_file(self) -> bool:
        return self.file_path is not None