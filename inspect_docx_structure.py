#!/usr/bin/env python3

import sys
import zipfile
from pathlib import Path


KEY_PHRASES = (
    "Performance Breaches",
    "Severity of Breach Identified",
    "Model(s) affected",
    "Model Risk Rating",
    "Model ID",
    "Model Name(s)",
)


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python3 inspect_docx_structure.py <file.docx>", file=sys.stderr)
        sys.exit(2)

    path = Path(sys.argv[1])
    with zipfile.ZipFile(str(path)) as archive:
        names = sorted(archive.namelist())
        word_xml = [name for name in names if name.startswith("word/") and name.endswith(".xml")]
        media = [name for name in names if name.startswith("word/media/")]

        print("DOCX:", path.name)
        print("word_xml_parts:", len(word_xml))
        print("media_files:", len(media))
        print()

        for name in word_xml:
            data = archive.read(name).decode("utf-8", errors="ignore")
            hits = [phrase for phrase in KEY_PHRASES if phrase.lower() in data.lower()]
            tbl_count = data.count("<w:tbl")
            txbx_count = data.count("txbxContent")
            if hits or tbl_count or txbx_count:
                print(name)
                print("  tables:", tbl_count)
                print("  textboxes:", txbx_count)
                if hits:
                    print("  hits:", ", ".join(hits))
                else:
                    print("  hits: none")


if __name__ == "__main__":
    main()
