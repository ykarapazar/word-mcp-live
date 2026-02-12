"""Default configuration values, overridable via environment variables."""

import os

DEFAULT_AUTHOR = os.environ.get("MCP_AUTHOR", "Author")
DEFAULT_INITIALS = os.environ.get("MCP_AUTHOR_INITIALS", "")
