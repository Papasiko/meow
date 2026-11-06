"""Package the app and prepared Windows 3.12 wheels into a transfer ZIP."""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT.parent / "OrgEvaluation_Windows10_Python312_OfflineReady.zip"
WHEELHOUSE = ROOT / "wheelhouse"

wheels = sorted(WHEELHOUSE.glob("*.whl")) if WHEELHOUSE.exists() else []
if len(wheels) < 8:
    print("ERROR: wheelhouse is not prepared yet.")
    print("Run prepare_offline_packages.bat on an Internet-connected Windows PC first.")
    print(f"Found {len(wheels)} wheel(s); at least 8 are expected.")
    sys.exit(1)

excluded_dirs = {".venv", "__pycache__", ".git"}
excluded_names = {"evaluation.db", "evaluation.db-shm", "evaluation.db-wal", ".secret_key"}
files: list[Path] = []
for path in ROOT.rglob("*"):
    if not path.is_file():
        continue
    rel = path.relative_to(ROOT)
    if any(part in excluded_dirs for part in rel.parts):
        continue
    if path.suffix == ".pyc" or path.name in excluded_names:
        continue
    if "backups" in rel.parts:
        continue
    files.append(path)

with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=7) as zf:
    for path in sorted(files):
        rel = path.relative_to(ROOT)
        arc = Path(ROOT.name) / rel
        zf.write(path, arc.as_posix())

with zipfile.ZipFile(OUT, "r") as zf:
    bad = zf.testzip()
    if bad:
        raise SystemExit(f"ZIP integrity check failed at: {bad}")
    names = zf.namelist()
    packaged_wheels = [n for n in names if n.startswith(f"{ROOT.name}/wheelhouse/") and n.endswith(".whl")]
    if len(packaged_wheels) < 8:
        raise SystemExit("ZIP integrity check failed: offline wheels are missing")
print(f"Created: {OUT}")
print(f"Size: {OUT.stat().st_size:,} bytes")
print(f"Files included: {len(files)}; offline wheels: {len(packaged_wheels)}")
