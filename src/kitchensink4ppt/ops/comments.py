"""Comments: modern threaded comments (write + reply + resolve + delete) and
legacy classic comments (read-only). No PowerPoint MCP server anywhere models
comments; python-pptx has no API at all (issue #487).

Two disjoint systems exist and do NOT interconvert (research Part VIII):

- MODERN (PowerPoint 2019/365, the default since 2021): namespace p188
  (.../powerpoint/2018/8/main). One authors part per package
  (ppt/authors.xml, p188:authorLst, BRACED-GUID author ids), one comments
  part per commented slide (p188:cmLst), related from the slide via the 2018
  reltype AND wired through the slide's p:extLst extension
  {6950BFC3-D8DA-4A85-94F7-54DA5524770B} carrying p188:commentRel. Forgetting
  the extension is the non-obvious half: the rel alone renders nothing.
  Replies are containment (p188:replyLst inside the parent p188:cm), not a
  parent-id link. Threads carry MS-PPTX task metadata; resolution is the
  status attribute on p188:cm (ST_CommentStatus: active/resolved/closed).
  PowerPoint 2019 and earlier cannot edit modern comments at all.
- LEGACY (classic, ECMA-376): p:cmLst per slide + one ppt/commentAuthors.xml
  (integer author ids). Identity is (authorId, idx); no replies, no resolved
  flag, plain p:text instead of a txBody. This module reads legacy comments
  but never writes them, and refuses to ADD modern comments to a deck that
  carries classic ones: Microsoft documents that the two systems do not mix
  in one file (a deck keeps whichever kind it has).

Anchoring: a modern comment anchors to the slide (pc:sldMkLst with the
slide's sldId + creationId) or to a shape (ac:deMkLst chaining doc + slide
monikers then ac:spMk with the shape id). The optional p188:pos offset is in
POINTS relative to the anchored object's top-left; callers pass EMU and this
module converts (EMU / 12700).

Ground truth: the corpus decks. military_brief.pptx carries real classic
comments (slide116); pmr_tables.pptx carries a real p188:authorLst, which is
the template for the author entries written here.

Timestamps are real UTC now (comments are review metadata, not corpus
content), formatted the way PowerPoint writes them: no timezone suffix,
millisecond precision.
"""

from __future__ import annotations

import os
import posixpath
import re
import uuid
from datetime import datetime, timedelta, timezone

from lxml import etree

from ..core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import (
    NSMAP,
    PRESENTATION_PART,
    PptxPackage,
    qn,
    rels_name,
    resolve_target,
)
from .read import resolve_slide, slide_table, slides_in_scope
from .shapes import _find_shape

# --------------------------------------------------------------- namespaces

NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
NS_AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
NS_A16 = "http://schemas.microsoft.com/office/drawing/2014/main"


def _q188(local: str) -> str:
    return f"{{{NS_P188}}}{local}"


def _qpc(local: str) -> str:
    return f"{{{NS_PC}}}{local}"


def _qac(local: str) -> str:
    return f"{{{NS_AC}}}{local}"


# ------------------------------------------------- reltypes / content types

RT_MODERN_COMMENTS = (
    "http://schemas.microsoft.com/office/2018/10/relationships/comments"
)
RT_MODERN_AUTHORS = (
    "http://schemas.microsoft.com/office/2018/10/relationships/authors"
)
RT_LEGACY_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
RT_LEGACY_AUTHORS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
)
CT_MODERN_COMMENTS = "application/vnd.ms-powerpoint.comments+xml"
CT_MODERN_AUTHORS = "application/vnd.ms-powerpoint.authors+xml"
CT_LEGACY_COMMENTS = (
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
)
CT_LEGACY_AUTHORS = (
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
)

MODERN_AUTHORS_PART = "ppt/authors.xml"

#: Slide extLst extension that makes the modern comments rel render.
EXT_URI_COMMENT_REL = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"
#: Slide creationId extension (p14:creationId val = the sldMk cId).
EXT_URI_SLIDE_CREATION_ID = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
#: Shape creationId extension (a16:creationId id = braced GUID).
EXT_URI_SHAPE_CREATION_ID = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"

ZERO_GUID = "{00000000-0000-0000-0000-000000000000}"

EMU_PER_POINT = 12700

_GUID_RE = re.compile(
    r"\{?([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}?\Z"
)


# ------------------------------------------------------------------ helpers


def default_author() -> str:
    """Server-level default author name: KS4P_COMMENT_AUTHOR env var when
    set, else "KitchenSink4PPT"."""
    return os.environ.get("KS4P_COMMENT_AUTHOR", "").strip() or "KitchenSink4PPT"


def _new_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _norm_guid(value, *, what: str = "comment_id") -> str:
    """'{...}' or bare GUID, any case -> canonical braced uppercase."""
    if not isinstance(value, str):
        raise PptMcpError(f"{what} must be a GUID string, got {value!r}")
    m = _GUID_RE.match(value.strip())
    if not m:
        raise PptMcpError(
            f"{what} {value!r} is not a GUID; use the comment_id returned by "
            "add_comment/list_comments"
        )
    return "{" + m.group(1).upper() + "}"


def _guid_or_none(value):
    if not isinstance(value, str):
        return None
    m = _GUID_RE.match(value.strip())
    return "{" + m.group(1).upper() + "}" if m else None


_last_stamp: list[str] = [""]


def _now_iso() -> str:
    """Real UTC now, PowerPoint's format, STRICTLY INCREASING within this
    process: PowerPoint orders replies by `created`, not document order, so
    two stamps landing in the same millisecond render in arbitrary order
    (found by the COM readback gate). Equal stamps are bumped by 1 ms."""
    dt = datetime.now(timezone.utc)
    stamp = f"{dt:%Y-%m-%dT%H:%M:%S}.{dt.microsecond // 1000:03d}"
    if stamp <= _last_stamp[0]:
        prev = datetime.strptime(
            _last_stamp[0], "%Y-%m-%dT%H:%M:%S.%f"
        ) + timedelta(milliseconds=1)
        stamp = f"{prev:%Y-%m-%dT%H:%M:%S}.{prev.microsecond // 1000:03d}"
    _last_stamp[0] = stamp
    return stamp


def _initials(name: str) -> str:
    parts = [w for w in re.split(r"[\s._-]+", name.strip()) if w]
    return ("".join(w[0] for w in parts[:3]).upper() or "?")


def _rel_target(src: str, dest: str) -> str:
    return posixpath.relpath(dest, posixpath.dirname(src)).replace("\\", "/")


def _require_text(text) -> str:
    if not isinstance(text, str) or not text.strip():
        raise PptMcpError("comment text must be a non-empty string")
    return text


def _txbody(text: str) -> etree._Element:
    """p188:txBody with one plain a:p per line."""
    body = etree.Element(_q188("txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    etree.SubElement(body, qn("a:lstStyle"))
    for line in text.split("\n"):
        p = etree.SubElement(body, qn("a:p"))
        if line:
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"))
            rpr.set("lang", "en-US")
            t = etree.SubElement(r, qn("a:t"))
            t.text = line
        else:
            etree.SubElement(p, qn("a:endParaRPr")).set("lang", "en-US")
    return body


def _txbody_text(body: etree._Element | None) -> str:
    """Plain text of an a:CT_TextBody (paragraphs by newline, a:br newline)."""
    if body is None:
        return ""
    paras = []
    for p in body.findall(qn("a:p")):
        chunks = []
        for node in p.iter():
            if node.tag == qn("a:t") and node.text:
                chunks.append(node.text)
            elif node.tag == qn("a:br"):
                chunks.append("\n")
        paras.append("".join(chunks))
    return "\n".join(paras)


# --------------------------------------------------------- part enumeration


def _rels_of_type(pkg: PptxPackage, part: str, reltype: str) -> list[tuple[str, str]]:
    """(rId, resolved target part) for every internal rel of one type."""
    try:
        rels = pkg.rels_for(part)
    except KeyError:
        return []
    out = []
    for rel in rels.getroot():
        if rel.get("Type") == reltype and rel.get("TargetMode") != "External":
            out.append(
                (rel.get("Id"), resolve_target(part, rel.get("Target", "")))
            )
    return out


def _modern_parts(pkg: PptxPackage, slide_part: str) -> list[tuple[str, str]]:
    return [
        (rid, part)
        for rid, part in _rels_of_type(pkg, slide_part, RT_MODERN_COMMENTS)
        if pkg.has_part(part)
    ]


def _legacy_parts(pkg: PptxPackage, slide_part: str) -> list[str]:
    return [
        part
        for _rid, part in _rels_of_type(pkg, slide_part, RT_LEGACY_COMMENTS)
        if pkg.has_part(part)
    ]


def _deck_has_legacy_comments(pkg: PptxPackage) -> bool:
    return any(
        _legacy_parts(pkg, rec["part"]) for rec in slide_table(pkg)
    )


# ------------------------------------------------------------ modern authors


def _modern_authors_part(pkg: PptxPackage) -> str | None:
    for _rid, part in _rels_of_type(pkg, PRESENTATION_PART, RT_MODERN_AUTHORS):
        if pkg.has_part(part):
            return part
    return None


def _ensure_modern_authors_part(pkg: PptxPackage) -> str:
    existing = _modern_authors_part(pkg)
    if existing is not None:
        return existing
    root = etree.Element(
        _q188("authorLst"),
        nsmap={"a": NSMAP["a"], "r": NSMAP["r"], "p188": NS_P188},
    )
    data = etree.tostring(
        etree.ElementTree(root), xml_declaration=True, encoding="UTF-8",
        standalone=True,
    )
    pkg.add_part_with_content_type(MODERN_AUTHORS_PART, data, CT_MODERN_AUTHORS)
    pkg.add_relationship(
        PRESENTATION_PART,
        RT_MODERN_AUTHORS,
        _rel_target(PRESENTATION_PART, MODERN_AUTHORS_PART),
    )
    return MODERN_AUTHORS_PART


def _author_guid(pkg: PptxPackage, name: str) -> str:
    """GUID of the p188:author entry for `name`: exact name match reuses the
    existing entry (author dedup), otherwise a new entry is created."""
    part = _ensure_modern_authors_part(pkg)
    root = pkg.root(part)
    for author in root.findall(_q188("author")):
        if author.get("name") == name:
            guid = _guid_or_none(author.get("id"))
            if guid:
                return guid
    guid = _new_guid()
    author = etree.SubElement(root, _q188("author"))
    author.set("id", guid)
    author.set("name", name)
    author.set("initials", _initials(name))
    author.set("userId", name)
    author.set("providerId", "None")
    pkg.mark_dirty(part)
    return guid


def _modern_author_map(pkg: PptxPackage) -> dict[str, dict]:
    out: dict[str, dict] = {}
    part = _modern_authors_part(pkg)
    if part is None:
        return out
    for author in pkg.root(part).findall(_q188("author")):
        guid = _guid_or_none(author.get("id"))
        if guid:
            out[guid] = {
                "name": author.get("name"),
                "initials": author.get("initials"),
            }
    return out


def _legacy_author_map(pkg: PptxPackage) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for _rid, part in _rels_of_type(pkg, PRESENTATION_PART, RT_LEGACY_AUTHORS):
        if not pkg.has_part(part):
            continue
        for author in pkg.root(part).findall(qn("p:cmAuthor")):
            out[author.get("id", "")] = {
                "name": author.get("name"),
                "initials": author.get("initials"),
            }
    return out


# ------------------------------------------------- modern part + slide wiring


def _wire_slide_ext(pkg: PptxPackage, slide_part: str, rid: str) -> None:
    """Ensure the slide's p:extLst carries the {6950BFC3-...} extension with
    p188:commentRel pointing at `rid`. The rel alone is NOT enough."""
    root = pkg.root(slide_part)
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        # p:extLst is the schema-last child of p:sld; appending is safe.
        ext_lst = etree.SubElement(root, qn("p:extLst"))
    for ext in ext_lst.findall(qn("p:ext")):
        if ext.get("uri") == EXT_URI_COMMENT_REL:
            rel = ext.find(_q188("commentRel"))
            if rel is not None and rel.get(qn("r:id")) == rid:
                return
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", EXT_URI_COMMENT_REL)
    rel = etree.SubElement(ext, _q188("commentRel"))
    rel.set(qn("r:id"), rid)
    pkg.mark_dirty(slide_part)


def _unwire_slide_ext(pkg: PptxPackage, slide_part: str, rid: str) -> None:
    root = pkg.root(slide_part)
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        return
    changed = False
    for ext in list(ext_lst.findall(qn("p:ext"))):
        if ext.get("uri") != EXT_URI_COMMENT_REL:
            continue
        rel = ext.find(_q188("commentRel"))
        if rel is not None and rel.get(qn("r:id")) == rid:
            ext_lst.remove(ext)
            changed = True
    if changed:
        if len(ext_lst) == 0:
            root.remove(ext_lst)
        pkg.mark_dirty(slide_part)


def _ensure_modern_part(pkg: PptxPackage, slide_part: str) -> tuple[str, str]:
    """(rId, comment part name) for the slide's modern comments part,
    creating part + content type + rel + extLst wiring when absent."""
    existing = _modern_parts(pkg, slide_part)
    if existing:
        rid, part = existing[0]
        _wire_slide_ext(pkg, slide_part, rid)  # heal a missing extension
        return rid, part
    part = pkg.next_partname("ppt/comments/modernComment_{}.xml")
    root = etree.Element(
        _q188("cmLst"),
        nsmap={
            "p188": NS_P188,
            "a": NSMAP["a"],
            "r": NSMAP["r"],
            "pc": NS_PC,
            "ac": NS_AC,
        },
    )
    data = etree.tostring(
        etree.ElementTree(root), xml_declaration=True, encoding="UTF-8",
        standalone=True,
    )
    pkg.add_part_with_content_type(part, data, CT_MODERN_COMMENTS)
    rid = pkg.add_relationship(
        slide_part, RT_MODERN_COMMENTS, _rel_target(slide_part, part)
    )
    _wire_slide_ext(pkg, slide_part, rid)
    return rid, part


# ------------------------------------------------------------------- anchors


def _slide_cid(pkg: PptxPackage, slide_part: str) -> str:
    """The slide's creationId val (p14:creationId in the slide extLst); "0"
    when the slide has none (synthetic decks). PowerPoint resolves the
    moniker through sldId; cId refines it."""
    root = pkg.root(slide_part)
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        return "0"
    for ext in ext_lst.findall(qn("p:ext")):
        if ext.get("uri") == EXT_URI_SLIDE_CREATION_ID:
            for child in ext:
                val = child.get("val") or child.get("id")
                if val:
                    return val
    return "0"


def _shape_creation_guid(shape: etree._Element) -> str:
    """The shape's a16:creationId GUID when present, else the zero GUID
    (PowerPoint accepts a zero creationId on shape monikers)."""
    for node in shape.iter():
        if node.tag == f"{{{NS_A16}}}creationId":
            guid = _guid_or_none(node.get("id") or node.get("val"))
            if guid:
                return guid
    return ZERO_GUID


def _slide_mk(pkg: PptxPackage, slide_part: str, slide_id: int, parent) -> None:
    etree.SubElement(parent, _qpc("docMk"))
    sld = etree.SubElement(parent, _qpc("sldMk"))
    sld.set("cId", _slide_cid(pkg, slide_part))
    sld.set("sldId", str(slide_id))


def _build_anchor(
    pkg: PptxPackage, slide_part: str, slide_id: int, anchor
) -> tuple[etree._Element, etree._Element | None, dict]:
    """(anchor element, optional p188:pos element, echo dict) from the tool
    anchor argument: None (slide, no offset), {"x": EMU, "y": EMU} (slide
    position), or {"shape_id": N[, "x": EMU, "y": EMU]} (shape anchor with
    optional offset from the shape's top-left)."""
    if anchor is None:
        anchor = {}
    if not isinstance(anchor, dict) or (
        set(anchor) - {"shape_id", "x", "y"}
    ):
        raise PptMcpError(
            f"invalid anchor {anchor!r}: use null, {{\"x\": EMU, \"y\": EMU}}"
            ", or {\"shape_id\": N}"
        )
    if ("x" in anchor) != ("y" in anchor):
        raise PptMcpError("anchor x and y must be given together (EMU)")

    pos = None
    echo: dict = {}
    if "x" in anchor:
        try:
            x_pt = round(int(anchor["x"]) / EMU_PER_POINT)
            y_pt = round(int(anchor["y"]) / EMU_PER_POINT)
        except (TypeError, ValueError):
            raise PptMcpError("anchor x/y must be integers in EMU")
        pos = etree.Element(_q188("pos"))
        pos.set("x", str(x_pt))
        pos.set("y", str(y_pt))
        echo.update({"x_emu": int(anchor["x"]), "y_emu": int(anchor["y"])})

    if "shape_id" in anchor:
        shape_id = anchor["shape_id"]
        if isinstance(shape_id, bool) or not isinstance(shape_id, int):
            raise PptMcpError(f"anchor shape_id must be an int, got {shape_id!r}")
        shape, _chain = _find_shape(pkg, slide_part, shape_id)  # validates
        mk = etree.Element(_qac("deMkLst"))
        _slide_mk(pkg, slide_part, slide_id, mk)
        local = etree.QName(shape).localname  # sp/grpSp/graphicFrame/cxnSp/pic
        sp = etree.SubElement(mk, _qac(local + "Mk"))
        sp.set("id", str(shape_id))
        sp.set("creationId", _shape_creation_guid(shape))
        echo.update({"type": "shape", "shape_id": shape_id})
        return mk, pos, echo

    mk = etree.Element(_qpc("sldMkLst"))
    _slide_mk(pkg, slide_part, slide_id, mk)
    echo["type"] = "slide"
    return mk, pos, echo


def _parse_anchor(cm: etree._Element) -> dict:
    """Anchor description of one modern p188:cm (its first moniker child)."""
    out: dict = {"type": "slide"}
    for child in cm:
        tag = etree.QName(child)
        if tag.namespace == NS_PC and tag.localname == "sldMkLst":
            break
        if tag.namespace == NS_AC and tag.localname == "deMkLst":
            for mk in child:
                mk_tag = etree.QName(mk)
                if mk_tag.namespace == NS_AC and mk_tag.localname.endswith("Mk"):
                    out["type"] = "shape"
                    try:
                        out["shape_id"] = int(mk.get("id", ""))
                    except ValueError:
                        pass
            break
        if tag.namespace == NS_AC and tag.localname == "txMkLst":
            out["type"] = "text"
            break
    pos = cm.find(_q188("pos"))
    if pos is not None:
        try:
            out["x_pt"] = int(pos.get("x", "0"))
            out["y_pt"] = int(pos.get("y", "0"))
        except ValueError:
            pass
    return out


# --------------------------------------------------------- comment location


def _find_cm(
    pkg: PptxPackage, slide_part: str, comment_id: str
) -> tuple[str, etree._Element, etree._Element | None]:
    """(part, p188:cm, reply-or-None) for a comment id on one slide. When
    the id names a reply, the returned cm is its thread root and the third
    member is the reply element itself."""
    target = _norm_guid(comment_id)
    known: list[str] = []
    for _rid, part in _modern_parts(pkg, slide_part):
        for cm in pkg.root(part).findall(_q188("cm")):
            cm_id = _guid_or_none(cm.get("id"))
            if cm_id:
                known.append(cm_id)
            if cm_id == target:
                return part, cm, None
            reply_lst = cm.find(_q188("replyLst"))
            if reply_lst is None:
                continue
            for reply in reply_lst.findall(_q188("reply")):
                if _guid_or_none(reply.get("id")) == target:
                    return part, cm, reply
    raise TargetNotFound(
        f"no modern comment {target} on that slide; thread ids present: "
        f"{known or 'none'}. Legacy classic comments are read-only here "
        "(list_comments shows them with legacy- ids)."
    )


# =============================================================== public API


def add_comment(
    pkg: PptxPackage,
    slide,
    text: str,
    author: str | None = None,
    anchor: dict | None = None,
) -> dict:
    """Add a modern threaded comment to a slide, creating the whole comment
    infrastructure from scratch when the deck has none: the p188 authors
    part (+ presentation rel + content type), the per-slide comments part
    (+ slide rel + content type), and the slide p:extLst commentRel wiring.
    `anchor`: None = the slide; {"x","y"} in EMU = a slide position;
    {"shape_id": N} = that shape. Author entries are deduplicated by exact
    name match; the default author is configurable (KS4P_COMMENT_AUTHOR).
    Refuses decks carrying classic comments (the two systems do not mix)."""
    text = _require_text(text)
    rec = resolve_slide(pkg, slide)
    if _deck_has_legacy_comments(pkg):
        raise UnsupportedStructure(
            "this deck carries classic (legacy) comments; PowerPoint does "
            "not mix classic and modern comments in one file and does not "
            "upgrade classic ones, so adding a modern comment here would "
            "produce a deck PowerPoint refuses to thread. Read them with "
            "list_comments; delete the classic comments in PowerPoint first "
            "if you want modern threads."
        )
    name = (author or default_author()).strip()
    if not name:
        raise PptMcpError("author must be a non-empty name")
    author_id = _author_guid(pkg, name)
    _rid, part = _ensure_modern_part(pkg, rec["part"])
    anchor_el, pos_el, anchor_echo = _build_anchor(
        pkg, rec["part"], rec["slide_id"], anchor
    )
    comment_id = _new_guid()
    created = _now_iso()
    cm = etree.SubElement(pkg.root(part), _q188("cm"))
    cm.set("id", comment_id)
    cm.set("authorId", author_id)
    cm.set("created", created)
    # Fixed child order: anchor, pos?, (replyLst), txBody.
    cm.append(anchor_el)
    if pos_el is not None:
        cm.append(pos_el)
    cm.append(_txbody(text))
    pkg.mark_dirty(part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "comment_id": comment_id,
        "author": name,
        "author_id": author_id,
        "created": created,
        "part": part,
        "anchor": anchor_echo,
    }


def reply_to_comment(
    pkg: PptxPackage,
    slide,
    comment_id: str,
    text: str,
    author: str | None = None,
) -> dict:
    """Append a threaded reply to a modern comment. Replies are containment
    (p188:reply inside the parent's p188:replyLst), so they need no anchor;
    replying to a reply is refused (threads are one level deep by design;
    reply to the thread root instead)."""
    text = _require_text(text)
    rec = resolve_slide(pkg, slide)
    part, cm, reply = _find_cm(pkg, rec["part"], comment_id)
    if reply is not None:
        raise PptMcpError(
            f"{_norm_guid(comment_id)} is a reply; threads are one level "
            f"deep. Reply to the thread root {_guid_or_none(cm.get('id'))} "
            "instead."
        )
    name = (author or default_author()).strip()
    if not name:
        raise PptMcpError("author must be a non-empty name")
    author_id = _author_guid(pkg, name)
    reply_lst = cm.find(_q188("replyLst"))
    if reply_lst is None:
        reply_lst = etree.Element(_q188("replyLst"))
        tx = cm.find(_q188("txBody"))
        if tx is not None:
            tx.addprevious(reply_lst)  # schema order: replyLst before txBody
        else:
            cm.append(reply_lst)
    reply_id = _new_guid()
    created = _now_iso()
    new = etree.SubElement(reply_lst, _q188("reply"))
    new.set("id", reply_id)
    new.set("authorId", author_id)
    new.set("created", created)
    new.append(_txbody(text))
    pkg.mark_dirty(part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "comment_id": _guid_or_none(cm.get("id")),
        "reply_id": reply_id,
        "author": name,
        "author_id": author_id,
        "created": created,
        "replies_in_thread": len(reply_lst),
    }


def _modern_record(cm: etree._Element, authors: dict, part: str) -> dict:
    author_id = _guid_or_none(cm.get("authorId"))
    who = authors.get(author_id, {})
    status = cm.get("status")
    replies = []
    reply_lst = cm.find(_q188("replyLst"))
    if reply_lst is not None:
        for reply in reply_lst.findall(_q188("reply")):
            r_author = _guid_or_none(reply.get("authorId"))
            r_who = authors.get(r_author, {})
            replies.append({
                "reply_id": _guid_or_none(reply.get("id")),
                "author": r_who.get("name"),
                "author_id": r_author,
                "created": reply.get("created"),
                "text": _txbody_text(reply.find(_q188("txBody"))),
            })
    return {
        "system": "modern",
        "comment_id": _guid_or_none(cm.get("id")),
        "author": who.get("name"),
        "author_initials": who.get("initials"),
        "author_id": author_id,
        "created": cm.get("created"),
        "text": _txbody_text(cm.find(_q188("txBody"))),
        "status": status,
        "resolved": (status or "").lower() == "resolved",
        "anchor": _parse_anchor(cm),
        "replies": replies,
        "part": part,
    }


def _legacy_record(cm: etree._Element, authors: dict, part: str) -> dict:
    author_id = cm.get("authorId", "")
    who = authors.get(author_id, {})
    pos = cm.find(qn("p:pos"))
    anchor = {"type": "position"}
    if pos is not None:
        anchor.update({
            "x_raw": pos.get("x"),
            "y_raw": pos.get("y"),
            "units": "raw p:pos values (legacy units are not EMU; unverified)",
        })
    text_el = cm.find(qn("p:text"))
    return {
        "system": "legacy",
        "comment_id": f"legacy-{author_id}-{cm.get('idx', '')}",
        "author": who.get("name"),
        "author_initials": who.get("initials"),
        "author_id": author_id,
        "created": cm.get("dt"),
        "text": text_el.text or "" if text_el is not None else "",
        "status": None,
        "resolved": None,  # classic comments have no resolved flag
        "anchor": anchor,
        "replies": [],  # classic comments cannot thread
        "part": part,
    }


def list_comments(pkg: PptxPackage, scope=None) -> dict:
    """Every comment in scope (None = all slides), BOTH systems: modern
    threaded comments with replies nested under their thread root and
    resolved status, and legacy classic comments (read-only; identity is
    (authorId, idx) rendered as legacy-A-I; no replies, no resolved flag)."""
    modern_authors = _modern_author_map(pkg)
    legacy_authors = _legacy_author_map(pkg)
    slides = []
    total = 0
    total_replies = 0
    for rec in slides_in_scope(pkg, scope):
        comments = []
        for _rid, part in _modern_parts(pkg, rec["part"]):
            for cm in pkg.root(part).findall(_q188("cm")):
                comments.append(_modern_record(cm, modern_authors, part))
        for part in _legacy_parts(pkg, rec["part"]):
            for cm in pkg.root(part).findall(qn("p:cm")):
                comments.append(_legacy_record(cm, legacy_authors, part))
        total += len(comments)
        total_replies += sum(len(c["replies"]) for c in comments)
        slides.append({
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
            "count": len(comments),
            "comments": comments,
        })
    return {
        "slides": slides,
        "total_comments": total,
        "total_replies": total_replies,
    }


def resolve_comment(
    pkg: PptxPackage, slide, comment_id: str, resolved: bool = True
) -> dict:
    """Set or clear a modern thread's resolved state via the status
    attribute MS-PPTX defines on p188:cm (active/resolved/closed).
    COMPATIBILITY: the flag rides the modern comment format, so PowerPoint
    2019 and earlier (which cannot read modern comments at all) never see
    it, and some 365 builds track resolution UI-side; the attribute is the
    documented interchange form. resolved=False removes the attribute
    (absence = active). Replies have no resolved state; resolve the root."""
    rec = resolve_slide(pkg, slide)
    part, cm, reply = _find_cm(pkg, rec["part"], comment_id)
    if reply is not None:
        raise PptMcpError(
            "replies have no resolved status; resolve the thread root "
            f"{_guid_or_none(cm.get('id'))}"
        )
    if resolved:
        cm.set("status", "resolved")
    elif "status" in cm.attrib:
        del cm.attrib["status"]
    pkg.mark_dirty(part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "comment_id": _guid_or_none(cm.get("id")),
        "resolved": bool(resolved),
    }


def delete_comment(
    pkg: PptxPackage, slide, comment_id: str, cascade_replies: bool = True
) -> dict:
    """Delete a modern comment thread, or a single reply when comment_id
    names a reply. A thread with replies requires cascade_replies=True
    (replies live INSIDE the thread root and cannot survive it). Deleting
    the last comment on a slide also removes the comments part, its rel,
    its content-type override, and the slide extLst wiring; author entries
    stay (other slides may reference them, matching PowerPoint)."""
    rec = resolve_slide(pkg, slide)
    part, cm, reply = _find_cm(pkg, rec["part"], comment_id)

    if reply is not None:
        reply_lst = reply.getparent()
        reply_lst.remove(reply)
        if len(reply_lst) == 0:
            cm.remove(reply_lst)
        pkg.mark_dirty(part)
        return {
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
            "deleted": _norm_guid(comment_id),
            "kind": "reply",
            "thread": _guid_or_none(cm.get("id")),
            "part_removed": False,
        }

    reply_lst = cm.find(_q188("replyLst"))
    n_replies = len(reply_lst) if reply_lst is not None else 0
    if n_replies and not cascade_replies:
        raise PptMcpError(
            f"thread {_norm_guid(comment_id)} has {n_replies} repl"
            f"{'y' if n_replies == 1 else 'ies'}; pass cascade_replies=true "
            "to delete the whole thread (replies cannot outlive their root)"
        )
    root = pkg.root(part)
    root.remove(cm)
    pkg.mark_dirty(part)

    part_removed = False
    if not root.findall(_q188("cm")):
        # Tear down the empty part: slide rel, extLst wiring, content type.
        slide_part = rec["part"]
        rels = pkg.rels_for(slide_part)
        for rel in list(rels.getroot()):
            if (
                rel.get("Type") == RT_MODERN_COMMENTS
                and rel.get("TargetMode") != "External"
                and resolve_target(slide_part, rel.get("Target", "")) == part
            ):
                _unwire_slide_ext(pkg, slide_part, rel.get("Id"))
                rels.getroot().remove(rel)
        pkg.mark_dirty(rels_name(slide_part))
        pkg.remove_part(part)
        part_rels = rels_name(part)
        if pkg.has_part(part_rels):
            pkg.remove_part(part_rels)
        pkg.remove_content_type_override(part)
        part_removed = True

    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "deleted": _norm_guid(comment_id),
        "kind": "thread",
        "replies_deleted": n_replies,
        "part_removed": part_removed,
    }


def comment_report(pkg: PptxPackage) -> dict:
    """Review-workflow rollup of the whole deck: every thread grouped by
    slide with authors, dates, resolved state, and nested replies, plus a
    markdown rendering ready to paste into review notes. Covers both
    modern threaded and legacy classic comments."""
    listing = list_comments(pkg, None)
    authors: set[str] = set()
    open_threads = 0
    resolved_threads = 0
    slides_with = []
    md: list[str] = [f"# Comment report — {pkg.path.name}", ""]
    for slide in listing["slides"]:
        if not slide["comments"]:
            continue
        slides_with.append(slide)
        md.append(f"## Slide {slide['slide_index'] + 1}")
        for c in slide["comments"]:
            who = c["author"] or "(unknown author)"
            authors.add(who)
            if c["resolved"]:
                resolved_threads += 1
            elif c["system"] == "modern":
                open_threads += 1
            flag = " [RESOLVED]" if c["resolved"] else ""
            tag = " (classic)" if c["system"] == "legacy" else ""
            when = c["created"] or "undated"
            text = " ".join((c["text"] or "").split())
            md.append(f"- **{who}** ({when}){flag}{tag}: {text}")
            for r in c["replies"]:
                r_who = r["author"] or "(unknown author)"
                authors.add(r_who)
                r_text = " ".join((r["text"] or "").split())
                md.append(
                    f"    - **{r_who}** ({r['created'] or 'undated'}): {r_text}"
                )
        md.append("")
    if not slides_with:
        md.append("No comments in this deck.")
    return {
        "file": str(pkg.path),
        "generated": _now_iso(),
        "total_comments": listing["total_comments"],
        "total_replies": listing["total_replies"],
        "open_threads": open_threads,
        "resolved_threads": resolved_threads,
        "authors": sorted(authors),
        "slides": slides_with,
        "markdown": "\n".join(md).rstrip() + "\n",
    }
