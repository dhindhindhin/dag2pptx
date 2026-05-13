"""
theme.py
----------
Design tokens for the DAG-diagram PowerPoint.

Colors, fonts, dimensions, and sizes are all centralized here.
The DAG "content" lives in the YAML; its "appearance" is managed here.

To adapt colors or sizes to in-house rules, edit only this file.
"""
from pptx.dml.color import RGBColor


# ==================== Font ====================
FONT = "Meiryo UI"


# ==================== Slide size (16:9) ====================
SLIDE_W_IN = 13.333
SLIDE_H_IN = 8


# ==================== Title ====================
# Uniform across all slides: white background + left-aligned navy text + underline.
TITLE_FONT_SIZE   = 20
TITLE_TEXT_COLOR  = RGBColor(0x1F, 0x3A, 0x5F)  # navy
TITLE_LINE_COLOR  = RGBColor(0x1F, 0x3A, 0x5F)
TITLE_LINE_WEIGHT = 2.0   # pt
TITLE_X           = 0.3   # left margin (inch)
TITLE_Y           = 0.22
TITLE_H           = 0.55
TITLE_LINE_Y      = 0.88  # y-coordinate of the title underline


# ==================== Content area ====================
TITLE_BAR_BOTTOM  = 0.80  # bottom edge of the title underline
CONTENT_PADDING   = 0.30


# ==================== Lead text (annotation right under the title underline) ====================
# Written under the `header:` key in the YAML. Placed just below the title underline as a lead.
# Plain text only, no frame or background. Tall enough to fit roughly 3 lines.
FOOTER_FONT_SIZE  = 11
FOOTER_TEXT       = RGBColor(0x3A, 0x3A, 0x3A)
FOOTER_X          = 0.3
FOOTER_Y          = 0.96  # title underline (0.88) + a small margin
FOOTER_H          = 0.70  # enough height for 3 lines


# ==================== Task box ====================
BOX_FILL           = RGBColor(0xDE, 0xEB, 0xF7)  # light blue
BOX_LINE           = RGBColor(0x2E, 0x75, 0xB6)  # blue border
BOX_TEXT           = RGBColor(0x1F, 0x3A, 0x5F)  # text
BOX_LINE_WEIGHT    = 1.25  # pt
BOX_LABEL_SIZE     = 10
BOX_SUBLABEL_SIZE  = 8
BOX_DELETED_FILL   = RGBColor(0xBF, 0xBF, 0xBF)  # gray for deleted boxes

# Reference dimensions used by auto-layout.
DEFAULT_BOX_W = 2.0  # baseline for font scaling
DEFAULT_BOX_H = 0.75


# ==================== Connector (elbow line) ====================
CONNECTOR_COLOR  = RGBColor(0x4A, 0x4A, 0x4A)
CONNECTOR_WEIGHT = 1.5  # pt


# ==================== Endpoint splitting for arrows ====================
# When multiple arrows leave (fan-out) or enter (fan-in) the same box, do not
# stack them all at the midpoint of the edge: divide the edge evenly and give
# each arrow its own attach point.
# Example: 2 arrows -> 1/3 and 2/3 positions; 3 arrows -> 1/4, 2/4, 3/4.
#
# Internally, the "vertical pairing" logic (lining up outer-outer / inner-inner
# pairs on the same x column for fan-out / fan-in) is always on. Combined with
# endpoint splitting, this keeps the vertical lanes symmetric and crossing-free.
ENDPOINT_SPLIT_ENABLED = True


# ==================== Status badges ====================
# Color intent:
#   blue  = new (informational)
#   green = changed (success / done)
#   red   = deleted (warning for destructive change)
BADGE_TEXT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
BADGE_FONT_SIZE  = 8
BADGE_H          = 0.26

BADGE_COLORS = {
    "new":     RGBColor(0x2E, 0x75, 0xB6),
    "changed": RGBColor(0x2E, 0x86, 0x3E),
    "deleted": RGBColor(0xC0, 0x00, 0x00),
}
BADGE_TEXTS = {
    "new":     "NEW",
    "changed": "CHANGED",
    "deleted": "DELETED",
}


# ==================== Labels (annotations) ====================
LABEL_DEFAULT_SIZE = 11
LABEL_CALLOUT_BG   = RGBColor(0xFD, 0xE8, 0xE8)
LABEL_CALLOUT_LINE = RGBColor(0xC0, 0x00, 0x00)


# ==================== Auto-layout tuning ====================
AUTO_MARGIN_X         = 0.5   # left/right slide margin (inch)
AUTO_COL_GAP_RATIO    = 0.12  # how much to shrink the box vs. column width (= inter-column gap)
AUTO_FONT_SCALE_MIN   = 0.78
AUTO_FONT_SCALE_MAX   = 1.20
AUTO_ROW_H_MULTIPLIER = 1.45  # row height = box height * this


# ==================== Named colors for YAML ====================
# Reference from YAML like `text_color: "badge_new"`.
NAMED_COLORS = {
    "title":         TITLE_TEXT_COLOR,
    "badge_new":     BADGE_COLORS["new"],
    "badge_changed": BADGE_COLORS["changed"],
    "badge_deleted": BADGE_COLORS["deleted"],
    "gray":          RGBColor(0xCC, 0xCC, 0xCC),
    "dark_gray":     RGBColor(0x55, 0x55, 0x55),
}
