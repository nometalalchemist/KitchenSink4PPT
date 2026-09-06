"""Sensitivity labels (Microsoft Purview / MIP) on a package.

A labeled deck carries `docMetadata/LabelInfo.xml`, related from the PACKAGE
ROOT rels (not from presentation.xml) with the classificationlabels
relationship type, plus a content-type override. The payload is one
`clbl:labelList` holding one `clbl:label` element whose id is the label GUID:

    <clbl:labelList xmlns:clbl="...mipLabelMetadata">
      <clbl:label id="{554eecc5-...}" enabled="1" method="Privileged"
                  siteId="{fae6d70f-...}" contentBits="0" removed="0"/>
    </clbl:labelList>

`enabled="0" removed="1"` is the tombstone PowerPoint writes when a user
CLEARS a label; it is not a classification, so read_label() returns None for
it and nothing carries it anywhere.

The label is metadata only. This module reads it, compares two of them, and
copies one package's label part into another verbatim; it never invents,
edits, or upgrades a classification, because the label text a user sees is
resolved from tenant policy that no file-tier tool can see.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import PptxPackage, resolve_target

LABEL_PART = "docMetadata/LabelInfo.xml"
RT_LABELS = (
    "http://schemas.microsoft.com/office/2020/02/relationships/"
    "classificationlabels"
)
CT_LABELS = "application/vnd.ms-office.classificationlabels+xml"
_CLBL = "{http://schemas.microsoft.com/office/2020/mipLabelMetadata}"


def label_part_of(pkg: PptxPackage) -> str | None:
    """The label part this package's ROOT rels point at, or None. The part
    name is read from the relationship rather than assumed, since only the
    relationship makes it a label."""
    try:
        rels = pkg.rels_for("")
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == RT_LABELS and rel.get("TargetMode") != "External":
            target = resolve_target("", rel.get("Target", ""))
            return target if pkg.has_part(target) else None
    return None


def read_label(pkg: PptxPackage) -> dict | None:
    """The active sensitivity label as {"id", "part", "method", "site_id",
    "enabled"}, or None when the deck carries no label (or only a cleared
    one). Malformed label XML reads as None rather than raising: a
    diagnostic must never be the thing that breaks a merge."""
    part = label_part_of(pkg)
    if part is None:
        return None
    try:
        root = etree.fromstring(pkg.raw_part(part))
    except etree.XMLSyntaxError:
        return None
    for el in root.iter(f"{_CLBL}label"):
        if el.get("removed") == "1" or el.get("enabled") == "0":
            continue
        label_id = el.get("id")
        if not label_id:
            continue
        out = {"id": label_id, "part": part, "enabled": True}
        for attr, key in (("method", "method"), ("siteId", "site_id")):
            if el.get(attr):
                out[key] = el.get(attr)
        return out
    return None


def copy_label(src: PptxPackage, dst: PptxPackage) -> str:
    """Copy src's label part into dst verbatim, with its content-type
    override and root relationship. The destination must not already carry a
    label part (callers check); returns the destination part name."""
    part = label_part_of(src)
    if part is None:  # pragma: no cover - callers check first
        raise ValueError("source package carries no label to copy")
    dst.add_part_with_content_type(LABEL_PART, src.raw_part(part), CT_LABELS)
    if label_part_of(dst) is None:
        dst.add_relationship("", RT_LABELS, LABEL_PART)
    return LABEL_PART


def describe(label: dict | None) -> str:
    """One-line human phrasing for warnings and refusals."""
    if label is None:
        return "no sensitivity label"
    method = label.get("method")
    return (
        f"sensitivity label {label['id']}"
        + (f" (method {method})" if method else "")
    )
