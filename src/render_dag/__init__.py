"""DAG -> PowerPoint rendering library."""

from .renderer import load_deck_from_yaml, render_deck, render_slide

__all__ = ["load_deck_from_yaml", "render_deck", "render_slide"]
