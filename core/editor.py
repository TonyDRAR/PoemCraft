import re
import unicodedata
import warnings
from functools import lru_cache


VOWELS = "aeiouyàâäéèêëîïôöùûüÿœæ"
VOWEL_GROUP_RE = re.compile(f"[{VOWELS}]+", re.IGNORECASE)
WORD_RE = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿœŒæÆ'’-]+")
ELISION_RE = re.compile(r"^[cdjlmnstqu]'(.+)$", re.IGNORECASE)
METRIC_RULES = {
    "Libre": None,
    "Alexandrin 12": [12],
    "Decasyllabe 10": [10],
    "Octosyllabe 8": [8],
    "Haiku 5/7/5": [5, 7, 5],
}
RHYME_LABELS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


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
    def get_metric_names() -> list[str]:
        return list(METRIC_RULES.keys())

    @staticmethod
    def get_line_metric_targets(metric_name: str, line_counts: list[int]) -> list[int | None]:
        rule = METRIC_RULES.get(metric_name)

        if not rule:
            return [None for _count in line_counts]

        targets = []
        verse_index = 0

        for count in line_counts:
            if count <= 0:
                targets.append(None)
                continue

            if len(rule) == 1:
                targets.append(rule[0])
                verse_index += 1
                continue

            targets.append(rule[verse_index] if verse_index < len(rule) else 0)
            verse_index += 1

        return targets

    @staticmethod
    def summarize_metric_progress(metric_name: str, line_counts: list[int]) -> tuple[int, int]:
        targets = Editor.get_line_metric_targets(metric_name, line_counts)
        checked_lines = [
            (count, target)
            for count, target in zip(line_counts, targets)
            if count > 0 and target is not None
        ]

        matching_lines = sum(1 for count, target in checked_lines if count == target)
        return matching_lines, len(checked_lines)

    @staticmethod
    def get_line_rhyme_labels(text: str) -> list[str]:
        rhyme_keys: dict[str, str] = {}
        next_label_index = 0
        labels = []

        for line in text.splitlines():
            key = Editor.get_line_rhyme_key(line)

            if not key:
                labels.append("")
                continue

            if key not in rhyme_keys:
                rhyme_keys[key] = Editor.build_rhyme_label(next_label_index)
                next_label_index += 1

            labels.append(rhyme_keys[key])

        return labels

    @staticmethod
    def get_rhyme_scheme(labels: list[str]) -> str:
        return " ".join(label for label in labels if label)

    @staticmethod
    def get_line_rhyme_key(line: str) -> str:
        words = WORD_RE.findall(line)

        if not words:
            return ""

        return Editor.get_word_rhyme_key(words[-1])

    @staticmethod
    def get_word_rhyme_key(word: str) -> str:
        word = Editor.remove_accents(Editor.normalize_word(word))

        if not word:
            return ""

        if len(word) > 3:
            word = re.sub(r"(?<=[^aeiouy])(?:e|es|ent)$", "", word)

        groups = list(VOWEL_GROUP_RE.finditer(word))

        if not groups:
            return word[-3:]

        start = groups[-1].start()

        if len(groups[-1].group()) == 1 and start > 0 and groups[-1].group() in "e":
            previous_groups = groups[:-1]

            if previous_groups:
                start = previous_groups[-1].start()

        return word[start:]

    @staticmethod
    def build_rhyme_label(index: int) -> str:
        label = ""
        index += 1

        while index:
            index, remainder = divmod(index - 1, len(RHYME_LABELS))
            label = RHYME_LABELS[remainder] + label

        return label

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
