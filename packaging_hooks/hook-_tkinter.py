from pathlib import Path
import sys


def _collect_tree(source: Path, dest_root: str):
    return [
        (str(path), str(Path(dest_root) / path.relative_to(source)))
        for path in source.rglob("*")
        if path.is_file()
    ]


_tcl_root = Path(sys.base_prefix) / "tcl"
datas = []

if (_tcl_root / "tcl8.6").is_dir():
    datas += _collect_tree(_tcl_root / "tcl8.6", "_tcl_data")

if (_tcl_root / "tk8.6").is_dir():
    datas += _collect_tree(_tcl_root / "tk8.6", "_tk_data")

if (_tcl_root / "tcl8").is_dir():
    datas += _collect_tree(_tcl_root / "tcl8", "tcl8")
