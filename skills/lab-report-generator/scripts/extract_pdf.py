#!/usr/bin/env python3
"""Extract paged PDF text; exit 2 if any page needs visual review."""

import argparse
from pathlib import Path
import sys

from pypdf import PdfReader


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.source.resolve() == args.output.resolve() or args.output.exists():
            raise ValueError("Output must be a new path different from the source.")
        reader = PdfReader(args.source)
        if reader.is_encrypted:
            raise ValueError("Encrypted PDF: ask the user for an accessible copy.")
        if not reader.pages:
            raise ValueError("PDF has no pages.")
        pages, review = [], []
        for number, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ""
            if len("".join(c for c in text if c.isalnum())) < 12:
                review.append(number)
                text = "[NEEDS_VISUAL_REVIEW: verify this page using its rendered image]\n" + text
            pages.append(f"=== Page {number} ===\n{text.strip()}\n")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with args.output.open("x", encoding="utf-8") as stream:
            stream.write("\n".join(pages))
        print(f"Extracted {len(pages)} pages; visual review required for pages: {review}")
        return 2 if review else 0
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
