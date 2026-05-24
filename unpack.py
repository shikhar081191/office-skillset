"""Unpack Office files (DOCX, PPTX, XLSX) for XML editing.

Extracts the ZIP archive and pretty-prints all XML files.
For DOCX files, optionally merges adjacent runs with identical formatting.

Usage:
    python unpack.py document.docx unpacked/
    python unpack.py presentation.pptx unpacked/
    python unpack.py spreadsheet.xlsx unpacked/
    python unpack.py document.docx unpacked/ --merge-runs false
"""

import argparse
import sys
import zipfile
from pathlib import Path

try:
    import defusedxml.minidom
except ImportError:
    print("ERROR: defusedxml not installed. Run: pip install defusedxml")
    sys.exit(1)

# Smart quotes → XML entities so they survive round-trips
SMART_QUOTE_REPLACEMENTS = {
    "\u201c": "&#x201C;",  # "
    "\u201d": "&#x201D;",  # "
    "\u2018": "&#x2018;",  # '
    "\u2019": "&#x2019;",  # '
}


def unpack(input_file: str, output_directory: str, merge_runs: bool = True) -> str:
    input_path = Path(input_file)
    output_path = Path(output_directory)
    suffix = input_path.suffix.lower()

    if not input_path.exists():
        return f"Error: {input_file} does not exist"

    if suffix not in {".docx", ".pptx", ".xlsx"}:
        return f"Error: {input_file} must be a .docx, .pptx, or .xlsx file"

    try:
        output_path.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(output_path)

        xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
        for xml_file in xml_files:
            _pretty_print_xml(xml_file)

        message = f"Unpacked {input_file} ({len(xml_files)} XML files) -> {output_directory}"

        if suffix == ".docx" and merge_runs:
            merge_count = _merge_runs_in_docx(output_path)
            message += f", merged {merge_count} runs"

        for xml_file in xml_files:
            _escape_smart_quotes(xml_file)

        return message

    except zipfile.BadZipFile:
        return f"Error: {input_file} is not a valid Office file (bad ZIP)"
    except Exception as e:
        return f"Error unpacking: {e}"


def _pretty_print_xml(xml_file: Path) -> None:
    try:
        content = xml_file.read_bytes()
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except Exception:
        pass  # Leave file as-is if XML is malformed


def _escape_smart_quotes(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_REPLACEMENTS.items():
            content = content.replace(char, entity)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        pass


def _merge_runs_in_docx(output_path: Path) -> int:
    """Merge adjacent runs with identical formatting in document.xml."""
    doc_xml = output_path / "word" / "document.xml"
    if not doc_xml.exists():
        return 0

    try:
        dom = defusedxml.minidom.parseString(doc_xml.read_text(encoding="utf-8"))
        root = dom.documentElement

        _remove_elements(root, "proofErr")
        _strip_rsid_attrs(root)

        runs = _find_elements(root, "r")
        containers = {run.parentNode for run in runs}

        merge_count = 0
        for container in containers:
            merge_count += _merge_runs_in_container(container)

        doc_xml.write_bytes(dom.toxml(encoding="UTF-8"))
        return merge_count

    except Exception:
        return 0


def _find_elements(root, tag: str) -> list:
    results = []
    def traverse(node):
        if node.nodeType == node.ELEMENT_NODE:
            name = node.localName or node.tagName
            if name == tag or name.endswith(f":{tag}"):
                results.append(node)
            for child in node.childNodes:
                traverse(child)
    traverse(root)
    return results


def _remove_elements(root, tag: str):
    for elem in _find_elements(root, tag):
        if elem.parentNode:
            elem.parentNode.removeChild(elem)


def _strip_rsid_attrs(root):
    for run in _find_elements(root, "r"):
        for attr in list(run.attributes.values()):
            if "rsid" in attr.name.lower():
                run.removeAttribute(attr.name)


def _get_child(parent, tag: str):
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            name = child.localName or child.tagName
            if name == tag or name.endswith(f":{tag}"):
                return child
    return None


def _is_run(node) -> bool:
    name = node.localName or node.tagName
    return name == "r" or name.endswith(":r")


def _next_element_sibling(node):
    sibling = node.nextSibling
    while sibling:
        if sibling.nodeType == sibling.ELEMENT_NODE:
            return sibling
        sibling = sibling.nextSibling
    return None


def _can_merge(run1, run2) -> bool:
    rpr1 = _get_child(run1, "rPr")
    rpr2 = _get_child(run2, "rPr")
    if (rpr1 is None) != (rpr2 is None):
        return False
    if rpr1 is None:
        return True
    return rpr1.toxml() == rpr2.toxml()


def _merge_runs_in_container(container) -> int:
    merge_count = 0
    children = list(container.childNodes)
    runs = [c for c in children if c.nodeType == c.ELEMENT_NODE and _is_run(c)]

    for run in runs:
        while True:
            next_elem = _next_element_sibling(run)
            if next_elem and _is_run(next_elem) and _can_merge(run, next_elem):
                # Move text content from next_elem into run
                for child in list(next_elem.childNodes):
                    if child.nodeType == child.ELEMENT_NODE:
                        name = child.localName or child.tagName
                        if name != "rPr" and not name.endswith(":rPr"):
                            run.appendChild(child)
                container.removeChild(next_elem)
                merge_count += 1
            else:
                break

    return merge_count


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Unpack an Office file for XML editing")
    parser.add_argument("input_file", help="Office file (.docx, .pptx, .xlsx)")
    parser.add_argument("output_directory", help="Directory to unpack into")
    parser.add_argument(
        "--merge-runs",
        type=lambda x: x.lower() == "true",
        default=True,
        metavar="true|false",
        help="Merge adjacent runs in DOCX (default: true)",
    )
    args = parser.parse_args()

    message = unpack(args.input_file, args.output_directory, merge_runs=args.merge_runs)
    print(message)
    sys.exit(1 if "Error" in message else 0)
