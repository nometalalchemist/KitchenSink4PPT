"""LaTeX equations as native PowerPoint math (OMML in DrawingML paragraphs).

Conversion pipeline ported verbatim from word-mcp ops/equations.py (both deps
MIT, pure Python, empirically Word-verified there): latex2mathml (LaTeX ->
MathML) then mathml2omml (MathML -> OMML). The produced OMML is parsed and
namespace-checked BEFORE anything in the package is touched, so a conversion
failure never modifies the deck. mathml2omml emits pure m:-namespace content
(m:r/m:rPr/m:sty/m:t), no Word (w:) elements, which is exactly what DrawingML
math accepts.

STORAGE FORMAT (the PPTX difference from .docx): PowerPoint stores math
inside a text-body paragraph via the Office 2010 drawing extension, wrapped
in markup compatibility:

    <a:p>
      <mc:AlternateContent xmlns:mc=".../markup-compatibility/2006">
        <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                   Requires="a14">
          <a14:m>
            <m:oMathPara xmlns:m=".../officeDocument/2006/math">
              <m:oMath>...</m:oMath>
            </m:oMathPara>
          </a14:m>
        </mc:Choice>
        <mc:Fallback>
          <a:r><a:rPr i="1"><a:latin typeface="Cambria Math"/></a:rPr>
               <a:t>linearized text</a:t></a:r>
        </mc:Fallback>
      </mc:AlternateContent>
    </a:p>

Ground truth (COM round-trip, PowerPoint 365, 2026-08-31): PowerPoint opens
decks carrying this wrapper clean, ingests the math into its text model
(TextRange.Text returns the linearized equation in Mathematical Alphanumeric
codepoints, proof the math engine parsed it), and re-serializes the same
mc:AlternateContent/mc:Choice[@Requires="a14"]/a14:m/m:oMathPara structure on
save — with no m:oMathParaPr (dropped from our earlier draft to match), with
Cambria Math a:rPr added to every m:r, and with the whole equation p:sp
additionally wrapped in a shape-level AlternateContent of its own. Those last
two are PowerPoint canonicalizations applied on ITS save; the minimal form
below is what it accepts. The mc:Fallback run is what pre-2010 consumers and
plain-text paths see.

READ-SIDE REALITY (by design, mirroring word-mcp): the read layer's
paragraph_text() walks a:r/a:fld/a:br/a:tab children only, so equations are
INVISIBLE to get_text / find_text / search_and_replace — which also means
search and replace can never corrupt an equation. The linearized approximation
is returned by this module's results (and by PowerPoint itself at render time).
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import NSMAP, PptxPackage, qn
from .read import resolve_slide

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"

_OMATH = f"{{{M_NS}}}oMath"
_OMATH_PARA = f"{{{M_NS}}}oMathPara"
_MT = f"{{{M_NS}}}t"
_A14_M = f"{{{A14_NS}}}m"

EMU_PER_INCH = 914400


class EquationConversionError(PptMcpError):
    """LaTeX did not convert to valid OMML; the presentation was not modified."""


def _rewrite_aligned(latex: str) -> str:
    # Upstream latex2mathml bug (carried from word-mcp): the `aligned`
    # environment emits malformed XML (unescaped entity). `align*` is
    # semantically equivalent for our purposes and converts cleanly.
    return latex.replace(r"\begin{aligned}", r"\begin{align*}").replace(
        r"\end{aligned}", r"\end{align*}"
    )


def _latex_to_omath(latex: str) -> etree._Element:
    """Convert LaTeX to a parsed m:oMath element, or raise (nothing mutated).

    mathml2omml emits a bare m:oMath string that uses the m: prefix without
    declaring the namespace, so we parse it inside a wrapper that declares
    xmlns:m — that single step both validates the XML and binds the prefix.
    """
    if not latex or not latex.strip():
        raise PptMcpError("latex must be non-empty")
    # File/system macros are meaningless inside an equation and would pass
    # through as literal text (word-mcp v1.5 adversarial F6) — refuse by name.
    _banned = re.search(
        r"\\(input|include|write\d*|openout|openin|read|immediate|usepackage)\b",
        latex,
    )
    if _banned:
        raise PptMcpError(
            f"\\{_banned.group(1)} is a file/preamble macro, not equation "
            "content; remove it — only math-mode LaTeX converts"
        )
    fixed = _rewrite_aligned(latex)
    try:
        import latex2mathml.converter as _l2m
        import mathml2omml as _m2o

        mathml = _l2m.convert(fixed, display="block")
        omml = _m2o.convert(mathml)
        wrapper = etree.fromstring(
            f'<wrap xmlns:m="{M_NS}">{omml}</wrap>'.encode("utf-8")
        )
    except Exception as exc:  # converters raise assorted types; all mean "bad input"
        raise EquationConversionError(
            f"could not convert LaTeX to PowerPoint math: {exc} "
            f"(input was: {latex!r}); the presentation was not changed"
        ) from exc
    children = list(wrapper)
    if len(children) != 1 or children[0].tag != _OMATH:
        raise EquationConversionError(
            f"converter produced unexpected OMML root for {latex!r}; "
            "the presentation was not changed"
        )
    omath = children[0]
    wrapper.remove(omath)
    return omath


def _approx_text(el: etree._Element) -> str:
    """Concatenated m:t content — a linear plain-text approximation only
    (structure like fraction bars and matrix layout is not represented)."""
    return "".join(t.text or "" for t in el.iter(_MT))


def _math_paragraph(omath: etree._Element, linear: str) -> etree._Element:
    """One a:p carrying the equation: mc:AlternateContent with the a14:m
    Choice (native math) and a plain italic Cambria Math run as Fallback."""
    p = etree.Element(qn("a:p"))
    ac = etree.SubElement(p, f"{{{MC_NS}}}AlternateContent", nsmap={"mc": MC_NS})
    choice = etree.SubElement(
        ac, f"{{{MC_NS}}}Choice", nsmap={"a14": A14_NS}
    )
    choice.set("Requires", "a14")
    a14m = etree.SubElement(choice, _A14_M)
    opara = etree.SubElement(a14m, _OMATH_PARA, nsmap={"m": M_NS})
    opara.append(omath)
    fallback = etree.SubElement(ac, f"{{{MC_NS}}}Fallback")
    r = etree.SubElement(fallback, qn("a:r"))
    rpr = etree.SubElement(r, qn("a:rPr"))
    rpr.set("i", "1")
    latin = etree.SubElement(rpr, qn("a:latin"))
    latin.set("typeface", "Cambria Math")
    t = etree.SubElement(r, qn("a:t"))
    t.text = linear
    return p


def insert_equation(
    pkg: PptxPackage,
    slide,
    latex: str,
    x,
    y,
    w=None,
    h=None,
) -> dict:
    """Insert a LaTeX equation as native PowerPoint math (editable in
    PowerPoint's equation editor, not an image) in a new text box at (x, y).

    Units follow the text-layer convention: floats and small ints are inches,
    ints of 10000+ are EMU. w/h default to 3.0 x 0.8 inches (the box autofits
    visually; PowerPoint centers the math per oMathParaPr). Any conversion
    failure raises with the LaTeX and the converter's message, and the
    presentation is not modified.

    NOTE: equations do not appear in get_text/find_text output (math lives
    outside the a:r text index by design); the returned "text" field carries
    the linearized approximation.
    """
    # Convert FIRST: any failure raises before the tree is touched.
    omath = _latex_to_omath(latex)
    linear = _approx_text(omath)

    from .text import _to_emu

    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    x_emu = _to_emu(x, None, "x")
    y_emu = _to_emu(y, None, "y")
    w_emu = _to_emu(w, None, "w") if w is not None else 3 * EMU_PER_INCH
    h_emu = _to_emu(h, None, "h") if h is not None else int(0.8 * EMU_PER_INCH)
    if w_emu <= 0 or h_emu <= 0:
        raise PptMcpError(
            f"equation box size must be positive; got w={w_emu} EMU, "
            f"h={h_emu} EMU"
        )
    from . import geometry as _g

    _g.check_emu_box(x_emu, y_emu, w_emu, h_emu, what="equation box")

    shape_id = pkg.next_shape_id(part)
    box_name = f"Equation {shape_id}"

    sp = etree.SubElement(sp_tree, qn("p:sp"))
    nvsp = etree.SubElement(sp, qn("p:nvSpPr"))
    cnv = etree.SubElement(nvsp, qn("p:cNvPr"))
    cnv.set("id", str(shape_id))
    cnv.set("name", box_name)
    cnvsp = etree.SubElement(nvsp, qn("p:cNvSpPr"))
    cnvsp.set("txBox", "1")
    etree.SubElement(nvsp, qn("p:nvPr"))

    sppr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(sppr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x_emu))
    off.set("y", str(y_emu))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(w_emu))
    ext.set("cy", str(h_emu))
    geom = etree.SubElement(sppr, qn("a:prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, qn("a:avLst"))
    etree.SubElement(sppr, qn("a:noFill"))

    body = etree.SubElement(sp, qn("p:txBody"))
    bodypr = etree.SubElement(body, qn("a:bodyPr"))
    bodypr.set("wrap", "square")
    bodypr.set("rtlCol", "0")
    etree.SubElement(body, qn("a:lstStyle"))
    body.append(_math_paragraph(omath, linear))

    pkg.mark_dirty(part)
    return {
        "equation_added": True,
        "latex": latex,
        "text": linear,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape_id,
        "name": box_name,
        "geometry": {"x": x_emu, "y": y_emu, "cx": w_emu, "cy": h_emu},
    }


def add_equation_to_shape(
    pkg: PptxPackage,
    slide,
    shape,
    latex: str,
    position: int | None = None,
) -> dict:
    """Append a LaTeX equation as a new math paragraph inside an existing
    shape's text body (existing paragraphs untouched).

    `shape` is a shape id (int) or shape name (str). `position` is the
    0-based paragraph index the equation paragraph is inserted BEFORE
    (None or past-the-end appends). Shapes without an editable text body
    (tables, pictures, charts, groups) refuse honestly. Any conversion
    failure raises before the shape is touched.
    """
    # Convert FIRST: any failure raises before the tree is touched.
    omath = _latex_to_omath(latex)
    linear = _approx_text(omath)

    from .text import _require_txbody, _resolve_shape

    rec = resolve_slide(pkg, slide)
    elem, kind = _resolve_shape(pkg, rec, shape)
    if kind in ("table", "chart", "picture", "group", "diagram", "ole"):
        raise UnsupportedStructure(
            f"shape {shape!r} on slide {rec['index']} is a {kind}; equations "
            "go into text-bearing shapes (text boxes, placeholders, "
            "autoshapes). Use insert_equation to place a new equation box."
        )
    body = _require_txbody(elem, kind, rec, create=True)
    paras = body.findall(qn("a:p"))
    new_p = _math_paragraph(omath, linear)
    if position is None or position >= len(paras):
        body.append(new_p)
        at = len(paras)
    else:
        if position < 0:
            raise PptMcpError(
                f"position must be a 0-based paragraph index, got {position}"
            )
        paras[position].addprevious(new_p)
        at = position

    pkg.mark_dirty(rec["part"])
    return {
        "equation_added": True,
        "latex": latex,
        "text": linear,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": int(elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}").get("id"))
        if elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}") is not None
        else None,
        "paragraph": at,
    }


def list_equations(pkg: PptxPackage, scope=None) -> dict:
    """Every native math equation on the selected slides (None = all).

    THIS IS THE READ PATH FOR MATH: equations never appear in get_text or
    find_text output (math runs are outside the a:r text index by design).
    Each entry carries slide index/id, shape id, the paragraph position of
    the equation within the shape, and the linearized text approximation.
    """
    from .read import iter_shapes, slides_in_scope

    entries: list[dict] = []
    for rec in slides_in_scope(pkg, scope):
        sp_tree = pkg.root(rec["part"]).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is None:
            continue
        for elem, kind, _z, _parent in iter_shapes(sp_tree):
            body = elem.find(qn("p:txBody"))
            if body is None:
                continue
            cnvpr = elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
            sid = int(cnvpr.get("id")) if cnvpr is not None else None
            for pi, p in enumerate(body.findall(qn("a:p"))):
                for om in p.iter(_OMATH):
                    entries.append(
                        {
                            "slide_index": rec["index"],
                            "slide_id": rec["slide_id"],
                            "shape_id": sid,
                            "paragraph": pi,
                            "text": _approx_text(om),
                        }
                    )
    return {"equations": entries, "equation_count": len(entries)}
