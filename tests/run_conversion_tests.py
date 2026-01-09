import os
from pathlib import Path
import sys

# ensure project root is on sys.path so we can import top-level modules when running this script
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from converter import md_to_docx, md_to_pptx, md_to_xlsx


def main():
    base = Path(__file__).parent
    md_file = base / 'sample.md'
    out_dir = base / 'out'
    out_dir.mkdir(exist_ok=True)

    md_text = md_file.read_text(encoding='utf-8')

    tests = [
        (md_to_docx, out_dir / 'sample.docx'),
        (md_to_pptx, out_dir / 'sample.pptx'),
        (md_to_xlsx, out_dir / 'sample.xlsx'),
    ]

    for fn, out in tests:
        try:
            fn(md_text, str(out))
            print(f"OK: wrote {out}")
        except Exception as e:
            print(f"ERROR running {fn.__name__}: {e}")


if __name__ == '__main__':
    main()