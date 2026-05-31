import unittest

from core.editor import Editor


class EditorRhymeTests(unittest.TestCase):
    def test_rhyme_key_groups_simple_rhymes(self):
        self.assertEqual(Editor.get_word_rhyme_key("amour"), Editor.get_word_rhyme_key("jour"))
        self.assertEqual(Editor.get_word_rhyme_key("vie"), Editor.get_word_rhyme_key("envie"))
        self.assertEqual(Editor.get_word_rhyme_key("lumiere"), Editor.get_word_rhyme_key("poussiere"))

    def test_rhyme_labels_keep_empty_lines_empty(self):
        text = "Le premier jour\n\nRevient l'amour"
        self.assertEqual(Editor.get_line_rhyme_labels(text), ["A", "", "A"])

    def test_rhyme_scheme_uses_repeated_labels(self):
        text = "\n".join(
            [
                "Je marche vers le jour",
                "La ville reprend vie",
                "Je cherche encore l'amour",
                "Une lumiere m'envie",
            ]
        )
        self.assertEqual(Editor.get_line_rhyme_labels(text), ["A", "B", "A", "B"])
        self.assertEqual(Editor.get_rhyme_scheme(["A", "B", "A", "B"]), "A B A B")

    def test_rhyme_labels_go_beyond_alphabet(self):
        self.assertEqual(Editor.build_rhyme_label(0), "A")
        self.assertEqual(Editor.build_rhyme_label(25), "Z")
        self.assertEqual(Editor.build_rhyme_label(26), "AA")


if __name__ == "__main__":
    unittest.main()
