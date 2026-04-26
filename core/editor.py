import re
import unicodedata
import warnings
from functools import lru_cache


VOWELS = "aeiouyàâäéèêëîïôöùûüÿœæ"
VOWEL_GROUP_RE = re.compile(f"[{VOWELS}]+", re.IGNORECASE)
WORD_RE = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿœŒæÆ'’-]+")
ELISION_RE = re.compile(r"^[cdjlmnstqu]'(.+)$", re.IGNORECASE)


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

    @staticmethod
    def count_syllables(text: str) -> int:
        return sum(Editor.count_line_syllables(text))

    @staticmethod
    def count_line_syllables(text: str) -> list[int]:
        return [Editor.count_verse_syllables(line) for line in text.splitlines()]

    @staticmethod
    def count_verse_syllables(line: str) -> int:
        words = WORD_RE.findall(line)
        total = 0

        for index, word in enumerate(words):
            next_word = words[index + 1] if index + 1 < len(words) else None
            total += Editor.count_word_syllables(word, next_word)

        return total

    @staticmethod
    def count_word_syllables(word: str, next_word: str | None = None) -> int:
        normalized_word = Editor.normalize_word(word)

        if not normalized_word:
            return 0

        lexical_info = Editor.lexical_syllable_info(normalized_word)

        if lexical_info is not None:
            lexical_count, has_elidable_final_e = lexical_info
            return lexical_count + Editor.poetic_final_e_bonus(
                normalized_word,
                next_word,
                has_elidable_final_e,
            )

        return Editor.fallback_syllable_count(normalized_word)

    @staticmethod
    def normalize_word(word: str) -> str:
        word = word.lower().replace("’", "'").strip("-'")
        match = ELISION_RE.match(word)

        if match:
            return match.group(1)

        return word

    @staticmethod
    def lexical_syllable_info(word: str) -> tuple[int, bool] | None:
        lexicon = Editor.get_lexicon()
        entries = lexicon.lexique.get(word)

        if not entries:
            entries = lexicon.lexique.get(Editor.remove_accents(word))

        if not entries:
            return None

        if not isinstance(entries, list):
            entries = [entries]

        best_entry = min((entry for entry in entries if entry.nbsyll), key=lambda entry: entry.nbsyll)
        orthographic_syllables = len(best_entry.orthosyll.split("-"))
        has_elidable_final_e = orthographic_syllables > best_entry.nbsyll

        return best_entry.nbsyll, has_elidable_final_e

    @staticmethod
    def poetic_final_e_bonus(word: str, next_word: str | None, has_elidable_final_e: bool) -> int:
        if not next_word or not has_elidable_final_e or not Editor.ends_with_mute_e(word):
            return 0

        next_word = Editor.normalize_word(next_word)

        if not next_word or Editor.starts_with_vowel_or_mute_h(next_word):
            return 0

        return 1

    @staticmethod
    def fallback_syllable_count(word: str) -> int:
        total = 0

        for part in re.split(r"[-']", word.lower()):
            part = part.strip()

            if len(part) <= 1 and part not in VOWELS:
                continue

            groups = VOWEL_GROUP_RE.findall(part)
            syllables = len(groups)

            if syllables > 1 and Editor.ends_with_mute_e(part):
                syllables -= 1

            total += max(1, syllables) if groups else 0

        return total

    @staticmethod
    def ends_with_mute_e(word: str) -> bool:
        return re.search(rf"[^{VOWELS}](e|es|ent)$", word, re.IGNORECASE) is not None

    @staticmethod
    def starts_with_vowel_or_mute_h(word: str) -> bool:
        return re.match(rf"^h?[{VOWELS}]", word, re.IGNORECASE) is not None

    @staticmethod
    def remove_accents(word: str) -> str:
        normalized = unicodedata.normalize("NFD", word)
        return "".join(char for char in normalized if unicodedata.category(char) != "Mn")

    @staticmethod
    @lru_cache(maxsize=1)
    def get_lexicon():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            from pylexique import Lexique383

            return Lexique383()
