#!/usr/bin/env python3
"""
render.py ~ CLI that generates a PowerPoint from a DAG-definition YAML.

Usage:
    python render.py INPUT.yaml [OUTPUT.pptx]

If OUTPUT is omitted, the .pptx is written next to INPUT.
"""
import sys
from pathlib import Path

# Add src/ to the path so the render_dag package can be imported.
sys.path.insert(0, str(Path(__file__).parent / "src"))
from render_dag import load_deck_from_yaml, render_deck


def main():
    if len(sys.argv) < 2:
        print("Usage: python render.py INPUT.yaml [OUTPUT.pptx]", file=sys.stderr)
        sys.exit(1)

    in_path = Path(sys.argv[1])
    if not in_path.exists():
        print(f"Error: file not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else in_path.with_suffix(".pptx")

    deck = load_deck_from_yaml(in_path)
    render_deck(deck, str(out_path))
    print(f"wrote {out_path} ({len(deck)} slides)")


if __name__ == "__main__":
    main()
