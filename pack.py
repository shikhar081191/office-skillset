"""Pack a directory back into a DOCX, PPTX, or XLSX file.

Condenses XML formatting and creates the Office ZIP archive.

Usage:
    python pack.py unpacked/ output.docx
    python pack.py unpacked/ output.pptx
    python pack.py unpacked/ output.xlsx
"""

import argparse
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    import defusedxml.minidom
except ImportError:
    print("ERROR: defusedxml not installed. Run: pip install defusedxml")
    sys.exit(1)


def pack(input_directory: str, output_file: str) -> str:
    input_dir = Path(input_directory)
    output_path = Path(output_file)
    suffix = output_path.suffix.lower()

    if not input_dir.is_dir():
        return f"Error: {input_dir} is not a directory"

    if suffix not in {".docx", ".pptx", ".xlsx"}:
        return f"Error: {output_file} must be a .docx, .pptx, or .xlsx file"

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_content_dir = Path(temp_dir) / "content"
            shutil.copytree(input_dir, temp_content_dir)

            # Condense XML (remove pretty-print whitespace)
            for pattern in ["*.xml", "*.rels"]:
                for xml_file in temp_content_dir.rglob(pattern):
                    _condense_xml(xml_file)

            output_path.parent.mkdir(parents=True, exist_ok=True)

            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in temp_content_dir.rglob("*"):
                    if f.is_file():
                        zf.write(f, f.relative_to(temp_content_dir))

        return f"Packed {input_directory} -> {output_file}"

    except Exception as e:
        return f"Error packing: {e}"


def _condense_xml(xml_file: Path) -> None:
    """Remove pretty-print whitespace from XML, preserving text content."""
    try:
        with open(xml_file, encoding="utf-8") as f:
            dom = defusedxml.minidom.parse(f)

        for element in dom.getElementsByTagName("*"):
            # Don't touch text runs — whitespace inside <w:t>, <a:t> etc. is significant
            if element.tagName.endswith(":t") or element.tagName == "t":
                continue

            for child in list(element.childNodes):
                if (
                    child.nodeType == child.TEXT_NODE
                    and child.nodeValue
                    and child.nodeValue.strip() == ""
                ) or child.nodeType == child.COMMENT_NODE:
                    element.removeChild(child)

        xml_file.write_bytes(dom.toxml(encoding="UTF-8"))

    except Exception as e:
        print(f"  Warning: Could not condense {xml_file.name}: {e}", file=sys.stderr)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pack a directory into an Office file")
    parser.add_argument("input_directory", help="Unpacked Office directory")
    parser.add_argument("output_file", help="Output file (.docx, .pptx, or .xlsx)")
    args = parser.parse_args()

    message = pack(args.input_directory, args.output_file)
    print(message)
    sys.exit(1 if "Error" in message else 0)
