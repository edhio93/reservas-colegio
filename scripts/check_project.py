from pathlib import Path
import py_compile
import sys

ROOT = Path(__file__).resolve().parents[1]
ignored = {"legacy"}
failures = []
count = 0

for path in ROOT.rglob("*.py"):
    if any(part in ignored for part in path.parts):
        continue
    count += 1
    try:
        py_compile.compile(str(path), doraise=True)
    except Exception as error:
        failures.append((path, error))

if failures:
    for path, error in failures:
        print(f"ERROR {path.relative_to(ROOT)}: {error}")
    sys.exit(1)

print(f"OK: {count} archivos Python compilaron correctamente.")
