"""Project-wide constants — single source of truth for magic numbers."""

# Page layout estimation (used by rendering.py, smart_features.py)
LINES_PER_PAGE = 45
CHARS_PER_LINE = 80

# LCS diff cap to keep O(n*m) memory/time acceptable
LCS_PARAGRAPH_CAP = 500

# EMU (English Metric Units) conversion factor
EMU_PER_INCH = 914400

# Rolling backup limit (.bak, .bak1, .bak2, .bak3)
MAX_BACKUP_COUNT = 3
BACKUP_SUFFIXES = [".bak"] + [f".bak{i}" for i in range(1, MAX_BACKUP_COUNT + 1)]
