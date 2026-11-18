"""masume (升目) — let AI agents read and fill Japanese "Excel 方眼紙" forms."""
__version__ = "0.1.0"

from .detect import Field, detect  # noqa: E402,F401
from .fill import FillReport, fill  # noqa: E402,F401
from .ooxml import load  # noqa: E402,F401
