import argparse
import sys
from pathlib import Path
from markitdown import MarkItDown


def main():
    parser = argparse.ArgumentParser(description="Convert a file to Markdown using markitdown")
    parser.add_argument("input", nargs="?", default="convocatoria-pias-2026.pdf", help="Input file (default: cap.pdf)")
    parser.add_argument("-o", "--output", help="Output .md file (default: same name as input with .md extension)")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix('.md')

    md = MarkItDown()
    try:
        result = md.convert(str(input_path))
    except Exception as e:
        print(f"Error converting {input_path}: {e}", file=sys.stderr)
        sys.exit(1)

    content = None
    # Prefer markdown-like attributes if available, fall back to text_content or str(result)
    for attr in ("markdown", "md", "text_content", "text"):
        content = getattr(result, attr, None)
        if content:
            break
    if content is None:
        content = str(result)

    output_path.write_text(content, encoding="utf-8")
    print(f"Wrote Markdown to {output_path}")


if __name__ == "__main__":
    main()