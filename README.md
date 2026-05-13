# Airflow DAG PowerPoint Generator

A tool that auto-generates explanatory PowerPoint slides for Cloud Composer /
Airflow DAGs from a YAML file. Write just the content (tasks and edges); the
appearance and layout are decided automatically.

---

## File layout

```
dag2pptx/
├── README.md                    # This file
├── requirements.txt
├── render.py                    # CLI entry point
├── src/
│   └── render_dag/              # Rendering library
│       ├── __init__.py
│       ├── renderer.py          # Rendering logic (auto layout, connectors)
│       └── theme.py             # Appearance (colors, fonts, dimensions)
├── examples/
│   ├── example_dag.yaml         # Starter YAML (a sample set of 9 slides)
│   └── example_dag.pptx         # The rendered output
└── blog/
    └── blog.mdx                 # Companion blog post
```

## Separation of concerns

| File                           | Role                                      | Touched              |
| ------------------------------ | ----------------------------------------- | -------------------- |
| `*.yaml`                       | **Content** (tasks, edges, descriptions)  | Every time you build a deck |
| `src/render_dag/theme.py`      | **Appearance** (colors, fonts, dimensions)| Once at the start    |
| `src/render_dag/renderer.py`   | **Rendering logic**                       | Rarely               |

## Install & usage

```bash
pip install -r requirements.txt
python render.py examples/example_dag.yaml output.pptx
```

Copy `examples/example_dag.yaml` as a starting point and rewrite its contents
to produce your own DAG diagram.

## Customizing colors and fonts

Edit `src/render_dag/theme.py`. For example:

```python
FONT = "Noto Sans JP"                    # Change the font

BADGE_COLORS = {                         # Badge colors
    "new":     RGBColor(0x00, 0x68, 0xB5),
    "changed": RGBColor(0x00, 0x80, 0x00),
    "deleted": RGBColor(0xC0, 0x00, 0x00),
}

TITLE_TEXT_COLOR = RGBColor(0x33, 0x33, 0x33)  # Title text color
TITLE_LINE_COLOR = RGBColor(0x33, 0x33, 0x33)  # Title underline color
```

## Slide structure

Each slide has 3 sections:

1. **Title bar** (top) — left-aligned navy title with an underline.
2. **Lead text** (right below the title) — about 3 lines summarizing the
   slide. Plain text without a frame. 11pt.
3. **DAG diagram** (the rest of the area) — boxes and arrows placed by
   auto layout.

## Writing `edges`

You can write edges with a string syntax that mirrors Airflow's `>>` operator.

```yaml
edges:
  - "start >> [extract_a, extract_b] >> transform >> [load, ml]"
  - "validate >> notify"
```

- `>>` chains left to right.
- `[a, b, c]` is fan-out / fan-in (as in Airflow, adjacent stages are joined
  by a Cartesian product).

## Auto layout

The slide's content determines the following automatically:

- **Column width**: slide width / max column count — the larger the DAG, the
  smaller the boxes.
- **Font**: proportional to box width (0.78x to 1.20x).
- **Box height**: tuned to the line count, with at least 0.18" of inter-row
  spacing.
- **Arrow endpoint splitting**: when multiple arrows leave or enter a single
  box, the edge is divided evenly so each arrow gets its own attach point.
- **Vertical pairing**: in fan-out / fan-in, upward and downward arrows are
  aligned on the same vertical lane for a symmetric, crossing-free layout.

## Key tunable toggles (`src/render_dag/theme.py`)

```python
ENDPOINT_SPLIT_ENABLED = True   # Split box edges so each arrow has its own
                                # attach point (recommended: True)
```

Vertical pairing is always on.

## PowerPoint placeholders that are stripped

The following placeholders that PowerPoint can otherwise show on open are
removed at render time:

- Date (`dt`)
- Footer (`ftr`)
- Slide number (`sldNum`)
