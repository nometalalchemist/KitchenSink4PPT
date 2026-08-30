"""PptxPackage: safe load/save layer for .pptx files.

Design guarantees:
- Parts that were never touched are written back byte-for-byte identical.
  A deck has dozens of parts (slides, layouts, masters, themes, notesSlides,
  media); an edit to slide 7 must not re-serialize slide 3.
- Saves are atomic: temp file -> validation -> os.replace. The original is
  never left half-written, and validation failure leaves it untouched.
- Auto-backup on by default: before each in-place save the current content is
  rotated into stable slots (prev.pptx / anchor.pptx) inside a hidden
  .ks4p-backups/ folder next to the presentation (see core.safesave).
- Files locked by PowerPoint are detected before any work happens.

PPTX structural rules baked in here (they differ from .docx and getting them
wrong is the classic corruption):
- Slide ORDER lives in presentation.xml's p:sldIdLst, resolved through the
  presentation rels. ZIP entry order and slideN.xml numbering mean nothing.
- EVERY part has its own .rels file; adding a part means adding its
  relationship AND its [Content_Types].xml override.
- Shape ids (p:cNvPr/@id) are unique PER SLIDE, never globally.
- p:sldId ids live in [256, 2147483647].
"""

from __future__ import annotations

import io
import os
import posixpath
import re
import zipfile
from pathlib import Path
from urllib.parse import unquote

from lxml import etree

from .errors import (
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    DocumentProtected,
    PptMcpError,
    ValidationFailed,
)
from .sandbox import check_path

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
}

PRESENTATION_PART = "ppt/presentation.xml"

#: Relationship type URIs used by the package layer.
RT_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)

#: Content types used by the package layer.
CT_SLIDE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)

#: p:sldId ids must live in this range (ECMA-376; python-pptx enforces the
#: same bounds).
SLIDE_ID_MIN = 256
SLIDE_ID_MAX = 2147483647

#: Schema-fixed order of p:presentation children. Inserting sldIdLst after
#: sldSz corrupts the file, so new children go in by rank, never appended.
_PRESENTATION_ORDER = (
    "p:sldMasterIdLst",
    "p:notesMasterIdLst",
    "p:handoutMasterIdLst",
    "p:sldIdLst",
    "p:sldSz",
    "p:notesSz",
    "p:smartTags",
    "p:embeddedFontLst",
    "p:custShowLst",
    "p:photoAlbum",
    "p:custDataLst",
    "p:kinsoku",
    "p:defaultTextStyle",
    "p:modifyVerifier",
    "p:extLst",
)


def qn(tag: str) -> str:
    """'p:sp' -> '{http://...}sp' (Clark notation)."""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


def _is_ole_encrypted(head: bytes) -> bool:
    # OLE compound file magic = password-protected or legacy .ppt, not a ZIP.
    return head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")


def rels_name(part: str) -> str:
    """'ppt/slides/slide1.xml' -> 'ppt/slides/_rels/slide1.xml.rels'."""
    d, fname = posixpath.split(part)
    return posixpath.join(d, "_rels", fname + ".rels") if d else posixpath.join("_rels", fname + ".rels")


def rels_source(rels_part: str) -> str:
    """'ppt/slides/_rels/slide1.xml.rels' -> 'ppt/slides/slide1.xml'.
    The package-root rels '_rels/.rels' maps to '' (the package itself)."""
    d, fname = posixpath.split(rels_part)
    parent = posixpath.dirname(d)
    source = fname[: -len(".rels")]
    if source == ".":  # '_rels/.rels'
        return ""
    return posixpath.join(parent, source) if parent else source


def resolve_target(source_part: str, target: str) -> str:
    """Resolve a relationship Target (part-relative, or absolute with a
    leading '/') against its source part's directory to a package part name."""
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target) if base else target)


class PptxPackage:
    """One .pptx opened for inspection or editing."""

    def __init__(self, path: str | os.PathLike):
        check_path(path, "open presentation")  # no-op unless KS4P_ALLOWED_ROOTS is set
        self.path = Path(path)
        self._raw: dict[str, bytes] = {}  # part name -> original bytes
        self._order: list[str] = []  # original entry order, preserved on save
        self._trees: dict[str, etree._ElementTree] = {}  # parsed parts
        self._dirty: set[str] = set()  # parts whose tree must be re-serialized
        self._load()

    # ---------- loading ----------

    def _load(self) -> None:
        if not self.path.exists():
            raise DocumentNotFound(f"No file at {self.path}")
        self._check_lock()
        head = self.path.read_bytes()[:8] if self.path.stat().st_size >= 8 else b""
        if _is_ole_encrypted(head):
            raise DocumentProtected(
                f"{self.path.name} is password-protected or a legacy binary "
                ".ppt; remove the password or convert to .pptx in PowerPoint "
                "first."
            )
        try:
            with zipfile.ZipFile(self.path) as zf:
                bad = zf.testzip()
                if bad is not None:
                    raise DocumentCorrupt(
                        f"{self.path.name}: corrupt ZIP entry '{bad}'."
                    )
                for info in zf.infolist():
                    self._raw[info.filename] = zf.read(info.filename)
                    self._order.append(info.filename)
        except zipfile.BadZipFile as exc:
            raise DocumentCorrupt(
                f"{self.path.name} is not a valid .pptx (bad ZIP): {exc}"
            ) from exc
        if PRESENTATION_PART not in self._raw:
            raise DocumentCorrupt(
                f"{self.path.name} has no {PRESENTATION_PART}; not a "
                "PowerPoint presentation."
            )

    def _check_lock(self) -> None:
        """PowerPoint holds an exclusive lock on open decks; detect it up front."""
        owner_file = self.path.with_name("~$" + self.path.name[-153:])
        try:
            with open(self.path, "r+b"):
                pass
        except PermissionError:
            hint = " (PowerPoint owner file present)" if owner_file.exists() else ""
            raise DocumentLocked(
                f"{self.path.name} is open in PowerPoint or locked by another "
                f"process{hint}. Close the file in PowerPoint and retry."
            ) from None

    # ---------- part access ----------

    def has_part(self, name: str) -> bool:
        return name in self._raw

    def part_names(self) -> list[str]:
        return list(self._order)

    def tree(self, name: str = PRESENTATION_PART) -> etree._ElementTree:
        """Parsed XML for a part. Parsing is cached; call mark_dirty() after edits."""
        if name not in self._trees:
            if name not in self._raw:
                raise KeyError(f"part not in package: {name}")
            parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
            self._trees[name] = etree.ElementTree(
                etree.fromstring(self._raw[name], parser=parser)
            )
        return self._trees[name]

    def root(self, name: str = PRESENTATION_PART) -> etree._Element:
        return self.tree(name).getroot()

    def presentation(self) -> etree._Element:
        """Root element of ppt/presentation.xml (the spine)."""
        return self.root(PRESENTATION_PART)

    def raw_part(self, name: str) -> bytes:
        return self._raw[name]

    def set_raw_part(self, name: str, data: bytes) -> None:
        """Add or replace a part with raw bytes (media, new XML parts)."""
        if name not in self._raw:
            self._order.append(name)
        self._raw[name] = data
        self._trees.pop(name, None)
        self._dirty.discard(name)  # raw bytes are authoritative now

    def mark_dirty(self, name: str = PRESENTATION_PART) -> None:
        if name not in self._trees:
            raise RuntimeError(f"mark_dirty before tree() for {name}")
        self._dirty.add(name)

    # ---------- relationships ----------

    def rels_for(self, part: str, *, create: bool = False) -> etree._ElementTree:
        """The .rels tree for a part ('' = the package root rels). With
        create=True a missing rels part is created empty; edits to the
        returned tree still require mark_dirty(rels_name(part))."""
        name = rels_name(part) if part else "_rels/.rels"
        if name not in self._raw:
            if not create:
                raise KeyError(f"no rels part for {part or '<package root>'}")
            empty = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                b'<Relationships xmlns='
                b'"http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )
            self.set_raw_part(name, empty)
        return self.tree(name)

    def next_rid(self, part: str) -> str:
        """Next free rId in a part's rels namespace. rIds are per source part
        and never assumed contiguous: max numeric suffix + 1."""
        try:
            rels = self.rels_for(part)
        except KeyError:
            return "rId1"
        highest = 0
        for rel in rels.getroot():
            m = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
            if m:
                highest = max(highest, int(m.group(1)))
        return f"rId{highest + 1}"

    def add_relationship(
        self,
        source_part: str,
        rel_type: str,
        target: str,
        *,
        external: bool = False,
    ) -> str:
        """Add a relationship from source_part; returns the new rId. `target`
        is part-relative for internal rels (as PowerPoint writes them) and a
        URI for external ones."""
        rid = self.next_rid(source_part)
        rels = self.rels_for(source_part, create=True)
        rel = etree.SubElement(rels.getroot(), f"{{{NSMAP['rel']}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")
        self.mark_dirty(rels_name(source_part) if source_part else "_rels/.rels")
        return rid

    def relationship_target(self, source_part: str, rid: str) -> str:
        """Resolve one rId of source_part to a package part name."""
        rels = self.rels_for(source_part)
        for rel in rels.getroot():
            if rel.get("Id") == rid:
                if rel.get("TargetMode") == "External":
                    raise PptMcpError(
                        f"{rid} in {source_part} is an external relationship"
                    )
                return resolve_target(source_part, rel.get("Target", ""))
        raise KeyError(f"{source_part} has no relationship {rid}")

    # ---------- content types ----------

    def add_content_type_override(self, part: str, content_type: str) -> None:
        """Ensure [Content_Types].xml carries an Override for `part`.
        Additive only: existing Overrides for other parts are never dropped."""
        ct_root = self.root("[Content_Types].xml")
        part_name = "/" + part
        for node in ct_root.findall(qn("ct:Override")):
            if node.get("PartName") == part_name:
                if node.get("ContentType") != content_type:
                    node.set("ContentType", content_type)
                    self.mark_dirty("[Content_Types].xml")
                return
        override = etree.SubElement(ct_root, qn("ct:Override"))
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        self.mark_dirty("[Content_Types].xml")

    def add_part_with_content_type(
        self, name: str, data: bytes, content_type: str
    ) -> None:
        """Add a part and its content-type override in one call. Forgetting
        the override is the #1 cause of PowerPoint repair prompts."""
        self.set_raw_part(name, data)
        self.add_content_type_override(name, content_type)

    # ---------- slides ----------

    def slide_parts(self) -> list[str]:
        """Slide part names in PRESENTATION ORDER: p:sldIdLst entries resolved
        through the presentation rels. NEVER ZIP order or filename numbering;
        slide2.xml can be the fifth slide after reordering."""
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            return []
        parts: list[str] = []
        for sld in sld_id_lst.findall(qn("p:sldId")):
            rid = sld.get(qn("r:id"))
            if rid is None:
                raise DocumentCorrupt("p:sldId entry without r:id")
            parts.append(self.relationship_target(PRESENTATION_PART, rid))
        return parts

    def next_shape_id(self, slide_part: str) -> int:
        """Next free shape id for one slide/layout/master/notesSlide part.
        Shape ids (p:cNvPr/@id) are unique PER PART, never package-global."""
        highest = 1  # the root spTree's own cNvPr is conventionally id=1
        for node in self.root(slide_part).iter(qn("p:cNvPr")):
            try:
                highest = max(highest, int(node.get("id", "0")))
            except ValueError:
                continue
        return highest + 1

    def _next_slide_id(self, sld_id_lst: etree._Element | None) -> int:
        highest = SLIDE_ID_MIN - 1
        if sld_id_lst is not None:
            for sld in sld_id_lst.findall(qn("p:sldId")):
                try:
                    highest = max(highest, int(sld.get("id", "0")))
                except ValueError:
                    continue
        if highest >= SLIDE_ID_MAX:
            raise PptMcpError("slide id space exhausted")
        return highest + 1

    def _insert_presentation_child(self, element: etree._Element) -> None:
        """Insert a child of p:presentation at its schema-fixed position."""
        pres = self.presentation()
        order = [qn(t) for t in _PRESENTATION_ORDER]
        try:
            rank = order.index(element.tag)
        except ValueError:
            raise PptMcpError(f"not a known p:presentation child: {element.tag}")
        for child in pres:
            if child.tag in order and order.index(child.tag) > rank:
                child.addprevious(element)
                return
        pres.append(element)

    def add_slide_part(
        self, slide_xml: bytes, *, layout_part: str
    ) -> dict:
        """Add a new slide to the deck ATOMICALLY in package terms: the part,
        its rels (with the layout relationship), the content-type override,
        the presentation relationship, and the p:sldId entry all land
        together, or an exception leaves the in-memory package unchanged.

        Adding a slide touches four registries; doing it half-way is the
        classic pptx corruption, which is why this is ONE method.
        """
        # Validate and precompute everything BEFORE mutating any structure.
        if layout_part not in self._raw:
            raise PptMcpError(f"layout part not in package: {layout_part}")
        try:
            etree.fromstring(slide_xml)
        except etree.XMLSyntaxError as exc:
            raise ValidationFailed(f"new slide XML is not well-formed: {exc}") from exc
        highest = 0
        for name in self._raw:
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
            if m:
                highest = max(highest, int(m.group(1)))
        new_part = f"ppt/slides/slide{highest + 1}.xml"
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        slide_id = self._next_slide_id(sld_id_lst)

        # Mutations begin; each step below is append-only and cannot fail
        # for structural reasons (all inputs were validated above).
        self.add_part_with_content_type(new_part, slide_xml, CT_SLIDE)
        layout_target = posixpath.relpath(layout_part, posixpath.dirname(new_part))
        layout_target = layout_target.replace(os.sep, "/")
        self.add_relationship(new_part, RT_SLIDE_LAYOUT, layout_target)
        pres_target = posixpath.relpath(new_part, posixpath.dirname(PRESENTATION_PART))
        pres_target = pres_target.replace(os.sep, "/")
        rid = self.add_relationship(PRESENTATION_PART, RT_SLIDE, pres_target)
        if sld_id_lst is None:
            sld_id_lst = etree.SubElement(pres, qn("p:sldIdLst"))
            pres.remove(sld_id_lst)
            self._insert_presentation_child(sld_id_lst)
        sld = etree.SubElement(sld_id_lst, qn("p:sldId"))
        sld.set("id", str(slide_id))
        sld.set(qn("r:id"), rid)
        self.mark_dirty(PRESENTATION_PART)
        return {"part": new_part, "slide_id": slide_id, "rid": rid}

    # ---------- saving ----------

    def _serialize(self, name: str) -> bytes:
        tree = self._trees[name]
        return etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def save(self, dest: str | os.PathLike | None = None, *, do_backup: bool = True) -> Path:
        """Atomic save. dest=None means save in place (with slot backup rotation
        by default: prev.pptx/anchor.pptx under .ks4p-backups/, see core.safesave).
        do_backup=False skips the rotation only; the atomic validated save always
        applies."""
        from . import safesave

        dest_path = Path(dest) if dest else self.path
        check_path(dest_path, "save presentation")
        in_place = dest_path.resolve() == self.path.resolve()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in self._order:
                data = self._serialize(name) if name in self._dirty else self._raw[name]
                zf.writestr(name, data)
        payload = buf.getvalue()

        # Validate the payload before touching the destination.
        self._validate_payload(payload)

        tmp = dest_path.with_name(dest_path.name + ".ppt-mcp-tmp")
        tmp.write_bytes(payload)
        try:
            # Capture the current (pre-mutation) content into the backup slots
            # via hardlink-then-replace: the presentation never leaves its own
            # path, and the old bytes survive the final replace below under
            # the prev slot's directory entry.
            if in_place and do_backup and dest_path.exists():
                try:
                    safesave.rotate_slots(dest_path)
                except PermissionError as exc:
                    raise PptMcpError(
                        f"cannot rotate backup slots for {dest_path.name}: "
                        f"another process holds a handle on a slot file. "
                        f"Close any program reading the backups and retry. "
                        f"({exc})"
                    ) from exc
            safesave.replace_with_retry(tmp, dest_path)
        except BaseException:
            tmp.unlink(missing_ok=True)
            raise
        # After a successful save the written bytes are the new baseline.
        if in_place:
            for name in self._dirty:
                self._raw[name] = self._serialize(name)
            self._dirty.clear()
        return dest_path

    @staticmethod
    def _validate_payload(payload: bytes) -> None:
        """Structural sanity before the destination is touched: valid ZIP,
        required roots present, every XML part well-formed, every internal
        relationship target resolvable, every sldIdLst entry resolvable.
        Dangling rels are the pptx equivalent of Word's repair prompt."""
        try:
            with zipfile.ZipFile(io.BytesIO(payload)) as zf:
                names = set(zf.namelist())
                if PRESENTATION_PART not in names:
                    raise ValidationFailed(f"output lost {PRESENTATION_PART}")
                if "[Content_Types].xml" not in names:
                    raise ValidationFailed("output lost [Content_Types].xml")
                roots: dict[str, etree._Element] = {}
                for name in names:
                    if name.endswith((".xml", ".rels")):
                        try:
                            roots[name] = etree.fromstring(zf.read(name))
                        except etree.XMLSyntaxError as exc:
                            raise ValidationFailed(
                                f"output part {name} is not well-formed XML: {exc}"
                            ) from exc

                # Every internal relationship target must exist in the package.
                rel_tag = f"{{{NSMAP['rel']}}}Relationship"
                pres_rid_targets: dict[str, str] = {}
                for name, root in roots.items():
                    if not name.endswith(".rels"):
                        continue
                    source = rels_source(name)
                    for rel in root.iter(rel_tag):
                        if rel.get("TargetMode") == "External":
                            continue
                        target = rel.get("Target", "")
                        resolved = resolve_target(source, target)
                        if resolved not in names:
                            # Partnames with escaped characters (%20) resolve
                            # through their unquoted spelling.
                            if unquote(resolved) in names:
                                resolved = unquote(resolved)
                            else:
                                raise ValidationFailed(
                                    f"dangling relationship in {name}: "
                                    f"{rel.get('Id')} -> {target} (no part "
                                    f"{resolved})"
                                )
                        if source == PRESENTATION_PART:
                            pres_rid_targets[rel.get("Id", "")] = resolved

                # Every slide in sldIdLst must resolve through those rels.
                pres = roots[PRESENTATION_PART]
                sld_id_lst = pres.find(qn("p:sldIdLst"))
                if sld_id_lst is not None:
                    for sld in sld_id_lst.findall(qn("p:sldId")):
                        rid = sld.get(qn("r:id"))
                        if rid not in pres_rid_targets:
                            raise ValidationFailed(
                                f"sldId {sld.get('id')} references {rid}, "
                                "which is not a presentation relationship"
                            )
        except zipfile.BadZipFile as exc:  # pragma: no cover
            raise ValidationFailed(f"output is not a valid ZIP: {exc}") from exc

    # ---------- Phase 2 additions (package-level primitives for slide CRUD) ----------

    def part_bytes(self, name: str) -> bytes:
        """Current effective bytes of a part: the serialized tree when the
        part is dirty, the original raw bytes otherwise. Use this instead of
        raw_part() when copying parts that may have in-memory edits."""
        if name in self._dirty:
            return self._serialize(name)
        return self._raw[name]

    def next_partname(self, template: str) -> str:
        """Collision-safe partname allocation. `template` carries one '{}'
        (e.g. 'ppt/slides/slide{}.xml'); scans the ACTUAL partnames in the
        package and returns template.format(max numeric suffix + 1). Never
        computed from a count: a deck that ever had deletions has holes, and
        len()+1 reuses an existing name (python-pptx's _next_slide_partname
        bug, corrupt package via duplicate ZIP entry names)."""
        prefix, _, suffix = template.partition("{}")
        pattern = re.compile(re.escape(prefix) + r"(\d+)" + re.escape(suffix) + r"\Z")
        highest = 0
        for name in self._raw:
            m = pattern.match(name)
            if m:
                highest = max(highest, int(m.group(1)))
        return template.format(highest + 1)

    def remove_part(self, name: str) -> None:
        """Drop a part from the in-memory package (raw bytes, cached tree,
        dirty flag, entry order). Content-type overrides and relationships
        pointing at the part are the caller's responsibility; the pre-save
        payload validation catches anything left dangling."""
        if name not in self._raw:
            raise KeyError(f"part not in package: {name}")
        del self._raw[name]
        self._order.remove(name)
        self._trees.pop(name, None)
        self._dirty.discard(name)

    def remove_content_type_override(self, part: str) -> bool:
        """Remove the [Content_Types].xml Override for `part` if present.
        Returns True when an Override was removed. Defaults are never touched
        (they cover other parts by extension)."""
        ct_root = self.root("[Content_Types].xml")
        part_name = "/" + part
        for node in ct_root.findall(qn("ct:Override")):
            if node.get("PartName") == part_name:
                ct_root.remove(node)
                self.mark_dirty("[Content_Types].xml")
                return True
        return False

    def register_slide_entry(
        self, slide_part: str, *, position: int | None = None
    ) -> dict:
        """Register an EXISTING slide part in the presentation spine: adds
        the presentation relationship and the p:sldId entry, at `position`
        (0-based index in deck order) or at the end. The slide part, its
        rels, and its content-type override must already be in the package;
        add_slide_part covers the build-from-XML path, this covers the
        clone-a-part path where the rels file is installed separately."""
        if slide_part not in self._raw:
            raise PptMcpError(f"slide part not in package: {slide_part}")
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        slide_id = self._next_slide_id(sld_id_lst)
        target = posixpath.relpath(slide_part, posixpath.dirname(PRESENTATION_PART))
        target = target.replace(os.sep, "/")
        rid = self.add_relationship(PRESENTATION_PART, RT_SLIDE, target)
        if sld_id_lst is None:
            sld_id_lst = etree.SubElement(pres, qn("p:sldIdLst"))
            pres.remove(sld_id_lst)
            self._insert_presentation_child(sld_id_lst)
        sld = etree.SubElement(sld_id_lst, qn("p:sldId"))
        sld.set("id", str(slide_id))
        sld.set(qn("r:id"), rid)
        entries = sld_id_lst.findall(qn("p:sldId"))
        index = len(entries) - 1
        if position is not None:
            if not 0 <= position < len(entries):
                raise PptMcpError(
                    f"position {position} out of range; the deck now has "
                    f"{len(entries)} slides (valid: 0..{len(entries) - 1})"
                )
            sld_id_lst.remove(sld)
            sld_id_lst.insert(position, sld)
            index = position
        self.mark_dirty(PRESENTATION_PART)
        return {"part": slide_part, "slide_id": slide_id, "rid": rid, "index": index}
