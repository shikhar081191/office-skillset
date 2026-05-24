"""Inject comments into an unpacked DOCX directory.

Must be run AFTER unpack.py. After adding comments, manually insert
<w:commentRangeStart>, <w:commentRangeEnd>, and <w:commentReference>
markers into word/document.xml at the desired location.

Usage:
    python comment.py unpacked/ 0 "Please review this section"
    python comment.py unpacked/ 1 "Reply text" --parent 0
    python comment.py unpacked/ 0 "Comment" --author "Shikhar"

After running, add these markers to document.xml around the target text:
    <w:commentRangeStart w:id="0"/>
    ... your target text ...
    <w:commentRangeEnd w:id="0"/>
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
"""

import argparse
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

try:
    import defusedxml.minidom
except ImportError:
    print("ERROR: defusedxml not installed. Run: pip install defusedxml")
    sys.exit(1)


NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def add_comment(
    unpacked_dir: str,
    comment_id: int,
    text: str,
    parent_id: int = None,
    author: str = "Claude",
    date: str = None,
) -> str:
    dir_path = Path(unpacked_dir)
    comments_path = dir_path / "word" / "comments.xml"
    rels_path = dir_path / "word" / "_rels" / "document.xml.rels"
    content_types_path = dir_path / "[Content_Types].xml"

    if not (dir_path / "word" / "document.xml").exists():
        return f"Error: {unpacked_dir} does not look like an unpacked DOCX"

    if date is None:
        date = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    _ensure_comments_relationship(rels_path)
    _ensure_comments_content_type(content_types_path)
    _add_comment_to_xml(comments_path, comment_id, text, author, date, parent_id)

    hint = (
        f"\nAdd these markers to word/document.xml around your target text:\n"
        f'  <w:commentRangeStart w:id="{comment_id}"/>\n'
        f"  ... target text ...\n"
        f'  <w:commentRangeEnd w:id="{comment_id}"/>\n'
        f'  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )
    return f"Added comment {comment_id} to {unpacked_dir}" + hint


def _ensure_comments_relationship(rels_path: Path):
    if not rels_path.exists():
        return

    content = rels_path.read_text(encoding="utf-8")
    if "comments.xml" in content:
        return

    rel_entry = (
        '<Relationship Id="rIdComments" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
        'Target="comments.xml"/>'
    )
    content = content.replace("</Relationships>", f"  {rel_entry}\n</Relationships>")
    rels_path.write_text(content, encoding="utf-8")


def _ensure_comments_content_type(content_types_path: Path):
    if not content_types_path.exists():
        return

    content = content_types_path.read_text(encoding="utf-8")
    if "comments" in content:
        return

    override = (
        '<Override PartName="/word/comments.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    )
    content = content.replace("</Types>", f"  {override}\n</Types>")
    content_types_path.write_text(content, encoding="utf-8")


def _add_comment_to_xml(comments_path: Path, comment_id: int, text: str, author: str, date: str, parent_id: int):
    w = W_NS

    if not comments_path.exists():
        # Create fresh comments.xml
        root_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
</w:comments>'''
        comments_path.write_text(root_xml, encoding="utf-8")

    content = comments_path.read_text(encoding="utf-8")

    # Build the comment XML string
    para_id = f"{(comment_id + 1) * 7:08X}"
    text_id = f"{(comment_id + 1) * 13:08X}"

    parent_attr = ""
    if parent_id is not None:
        parent_attr = f' w15:paraIdParent="{(parent_id + 1) * 7:08X}"'

    comment_xml = (
        f'  <w:comment w:id="{comment_id}" w:author="{author}" w:date="{date}" w:initials="{author[:2].upper()}">\n'
        f'    <w:p w14:paraId="{para_id}" w14:textId="{text_id}"{parent_attr}>\n'
        f'      <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>\n'
        f'      <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:annotationRef/></w:r>\n'
        f'      <w:r><w:t>{text}</w:t></w:r>\n'
        f'    </w:p>\n'
        f'  </w:comment>\n'
    )

    content = content.replace("</w:comments>", comment_xml + "</w:comments>")
    comments_path.write_text(content, encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Add a comment to an unpacked DOCX")
    parser.add_argument("unpacked_dir", help="Unpacked DOCX directory")
    parser.add_argument("comment_id", type=int, help="Comment ID (integer)")
    parser.add_argument("text", help="Comment text")
    parser.add_argument("--parent", type=int, default=None, help="Parent comment ID (for replies)")
    parser.add_argument("--author", default="Claude", help="Comment author (default: Claude)")
    args = parser.parse_args()

    result = add_comment(args.unpacked_dir, args.comment_id, args.text, args.parent, args.author)
    print(result)
    sys.exit(1 if "Error" in result else 0)
