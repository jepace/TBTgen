"""Stamps a new version on tbtgen_sheet.gs (first line + SCRIPT_VERSION) and build.SCRIPT_VERSION.
Run after every change to tbtgen_sheet.gs:  python3 bump_script_version.py"""
import datetime as dt
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
v = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
gs, py = ROOT / "tbtgen_sheet.gs", ROOT / "build.py"
s = gs.read_text()
s, a = re.subn(r"^// ===== TBTgen sheet script · VERSION [^·]+·", f"// ===== TBTgen sheet script · VERSION {v} ·", s, count=1)
s, b = re.subn(r"const SCRIPT_VERSION = '[^']*';", f"const SCRIPT_VERSION = '{v}';", s, count=1)
t, c = re.subn(r'^SCRIPT_VERSION = "[^"]*"', f'SCRIPT_VERSION = "{v}"', py.read_text(), count=1, flags=re.M)
assert a == b == c == 1, "version markers not found"
gs.write_text(s)
py.write_text(t)
print("sheet script version:", v)
