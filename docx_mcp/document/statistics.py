"""Document statistics mixin: word count, character count, paragraph/table/image/section counts."""

from __future__ import annotations

from .base import WP, W


class StatisticsMixin:

    def get_word_count(self) -> int:
        root = self._tree("word/document.xml")
        if root is None:
            return 0
        body = root.find(f"{W}body")
        if body is None:
            return 0
        return sum(
            len(t.text.split()) for t in body.iter(f"{W}t") if t.text
        )

    def get_reading_time(self, words_per_minute: int = 200) -> dict:
        """Estimate reading time based on word count.

        Args:
            words_per_minute: Reading speed (default 200 wpm).

        Returns:
            {"word_count": int, "words_per_minute": int, "minutes": float, "seconds": int}
        """
        if words_per_minute <= 0:
            raise ValueError("words_per_minute must be a positive integer")
        word_count = self.get_word_count()
        minutes = word_count / words_per_minute
        return {
            "word_count": word_count,
            "words_per_minute": words_per_minute,
            "minutes": round(minutes, 1),
            "seconds": round(minutes * 60),
        }

    def get_statistics(self) -> dict:
        root = self._tree("word/document.xml")
        if root is None:
            return {
                "word_count": 0,
                "character_count": 0,
                "paragraph_count": 0,
                "table_count": 0,
                "image_count": 0,
                "section_count": 1,
            }
        body = root.find(f"{W}body")
        if body is None:
            return {
                "word_count": 0,
                "character_count": 0,
                "paragraph_count": 0,
                "table_count": 0,
                "image_count": 0,
                "section_count": 1,
            }

        texts = [t.text for t in body.iter(f"{W}t") if t.text]
        word_count = sum(len(t.split()) for t in texts)
        character_count = sum(len(t) for t in texts)

        paragraph_count = sum(
            1 for child in body if child.tag == f"{W}p"
        )
        table_count = sum(
            1 for child in body if child.tag == f"{W}tbl"
        )
        image_count = sum(
            1 for _ in body.iter(f"{WP}inline")
        ) + sum(
            1 for _ in body.iter(f"{WP}anchor")
        )

        explicit_sect_pr = sum(
            1 for p in body
            if p.tag == f"{W}p"
            for ppr in [p.find(f"{W}pPr")]
            if ppr is not None and ppr.find(f"{W}sectPr") is not None
        )
        section_count = explicit_sect_pr + 1

        return {
            "word_count": word_count,
            "character_count": character_count,
            "paragraph_count": paragraph_count,
            "table_count": table_count,
            "image_count": image_count,
            "section_count": section_count,
        }
