"""USAF Operating Instruction auto-formatter.

Rewrites a .docx to comply with AFH 33-337 (Tongue and Quill) and
DAFMAN 90-161 (Publishing Processes and Procedures).
"""

from .meta import OIMeta
from .profile import FormattingProfile

__all__ = ["FormattingProfile", "OIMeta", "__version__"]
__version__ = "0.2.0"
