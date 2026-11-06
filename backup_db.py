from datetime import datetime
import os
import sqlite3
from pathlib import Path

base = Path(__file__).resolve().parent
source = Path(os.environ.get("EVALUATION_DB_PATH", str(base / "instance" / "evaluation.db")))
dest_dir = base / "backups"
dest_dir.mkdir(exist_ok=True)
if not source.exists():
    raise SystemExit(f"Database not found: {source}")
target = dest_dir / f"evaluation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
with sqlite3.connect(source) as src, sqlite3.connect(target) as dst:
    src.backup(dst)
print(f"Database backup created: {target}")
