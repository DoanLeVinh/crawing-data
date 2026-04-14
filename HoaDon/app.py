from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from web.app import app  # noqa: E402


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
