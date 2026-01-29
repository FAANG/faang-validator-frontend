#!/usr/bin/env python3
"""
Parse ENA (European Nucleotide Archive) submission response XML.

Usage:
  - As module: result = parse_ena_response(xml_bytes_or_str)
  - From file:  python parse_ena_response.py path/to/receipt.xml
  - From stdin: curl ... | python parse_ena_response.py -

Returns a string representation of the parsed result.
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.etree import ElementTree as ET


def parse_ena_response(xml_input: bytes | str) -> str:
    """
    Parse ENA RECEIPT XML and return a human-readable string.

    Args:
        xml_input: Raw ENA response as bytes or str (e.g. from curl stdout).

    Returns:
        Multi-line string with success status, errors, info messages,
        and any SUBMISSION/ANALYSIS accessions.
    """
    if isinstance(xml_input, bytes):
        text = xml_input.decode("utf-8")
    else:
        text = xml_input

    # Access Denied (plain text, not XML)
    if "Access Denied" in text:
        return "success: false\nerrors:\n  - Access Denied\ninfo: []"

    try:
        root = ET.fromstring(xml_input)
    except ET.ParseError as e:
        return f"parse_error: {e}\nraw_preview: {text[:500]!r}"

    return _format_result(_parse_root(root))


def _parse_root(root: ET.Element) -> dict:
    """Extract structured data from RECEIPT root."""
    errors = []
    info = []
    actions = []

    for messages in root.findall("MESSAGES"):
        for e in messages.findall("ERROR"):
            t = e.text
            errors.append(t.strip() if t else "")
        for i in messages.findall("INFO"):
            t = i.text
            info.append(t.strip() if t else "")

    success = len(errors) == 0
    receipt_success = root.get("success", "").lower() == "true"

    # ACTIONS element (e.g. <ACTIONS>ADD</ACTIONS>)
    for a in root.findall("ACTIONS"):
        t = a.text
        if t:
            actions.append(t.strip())

    # Optional: SUBMISSION / ANALYSIS accessions
    submissions = []
    for s in root.findall("SUBMISSION"):
        submissions.append(
            {"alias": s.get("alias", ""), "accession": s.get("accession", "")}
        )
    analyses = []
    for a in root.findall("ANALYSIS"):
        analyses.append(
            {"alias": a.get("alias", ""), "accession": a.get("accession", "")}
        )

    return {
        "success": success,
        "receipt_success": receipt_success,
        "errors": errors,
        "info": info,
        "actions": actions,
        "submissions": submissions,
        "analyses": analyses,
    }


def _format_result(data: dict) -> str:
    """Turn parsed dict into a string."""
    lines = [
        f"success: {str(data['success']).lower()}",
        f"receipt_success: {str(data['receipt_success']).lower()}",
        "",
        "errors:",
    ]
    if data["errors"]:
        for e in data["errors"]:
            lines.append(f"  - {e}")
    else:
        lines.append("  (none)")
    lines.append("")
    lines.append("info:")
    if data["info"]:
        for i in data["info"]:
            lines.append(f"  - {i}")
    else:
        lines.append("  (none)")

    if data.get("actions"):
        lines.append("")
        lines.append("actions:")
        for a in data["actions"]:
            lines.append(f"  - {a}")

    if data["submissions"]:
        lines.append("")
        lines.append("submissions:")
        for s in data["submissions"]:
            lines.append(f"  - alias: {s['alias']}, accession: {s['accession']}")

    if data["analyses"]:
        lines.append("")
        lines.append("analyses:")
        for a in data["analyses"]:
            lines.append(f"  - alias: {a['alias']}, accession: {a['accession']}")

    return "\n".join(lines)


def main() -> None:
    """
    Simple CLI:

    - With argument: python ena-response-parsing.py <file.xml|->
    - With no argument: if ./response.xml exists, it will be used.
    """
    if len(sys.argv) < 2:
        default = Path("response.xml")
        if default.is_file():
            path = str(default)
        else:
            print(
                "Usage: python ena-response-parsing.py <file.xml|->",
                file=sys.stderr,
            )
            print(
                "  -  Read from file, or '-' for stdin.", file=sys.stderr
            )
            sys.exit(1)
    else:
        path = sys.argv[1]

    if path == "-":
        raw = sys.stdin.buffer.read()
    else:
        with open(path, "rb") as f:
            raw = f.read()

    result = parse_ena_response(raw)
    print(result)


if __name__ == "__main__":
    main()
