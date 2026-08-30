"""Workflow guidance: recommended tool sequences for common multi-step tasks.

Discoverability rule 4 (PLAN section 2): get_workflows ships in lite and every
recipe NAMES the packs it needs, so a fresh session learns both the sequence
and the exact enable_tools call before touching a deck. Every tool named in a
step MUST be a registered tool on the server (tests assert this against the
live registry), so a workflow can never point at a tool that does not exist.

The sequences are recommendations, not scripts: steps marked optional can be
skipped, and the notes carry the judgment calls a caller should know about.
"""

from __future__ import annotations

from ..core.errors import PptMcpError

# task -> {summary, packs, steps: [{tool, why, optional?}], notes: [...]}
WORKFLOWS: dict[str, dict] = {
    "build-a-diagram": {
        "summary": (
            "Turn an SVG (or a from-scratch shape layout) into native, "
            "individually editable PowerPoint shapes with glued connectors, "
            "then verify by rendering."
        ),
        "packs": ["graphics", "assembly-export"],
        "steps": [
            {"tool": "enable_tools",
             "why": "packs=['graphics','assembly-export'] turns on the shape "
                    "engine and the render-to-verify loop"},
            {"tool": "get_presentation_view",
             "why": "see the target slide's current shapes and anchors first"},
            {"tool": "svg_to_shapes",
             "why": "compile the SVG into one grouped tree of native shapes "
                    "(no rasterization); returns every created shape id"},
            {"tool": "insert_connector", "optional": True,
             "why": "add glued arrows between shapes by id; glued ends follow "
                    "the shape when it moves"},
            {"tool": "export_slide_image",
             "why": "render the slide to PNG and LOOK at it; layout bugs are "
                    "visual, not structural"},
            {"tool": "set_shape",
             "why": "nudge position, size, fill, or text by shape id and "
                    "re-render until it is right"},
        ],
        "notes": [
            "insert_shape builds presets and freeform geometry directly when "
            "there is no SVG to compile.",
            "align_shapes and distribute_shapes do mechanical cleanup that "
            "eyeballing coordinates cannot.",
        ],
    },
    "build-a-table-report": {
        "summary": (
            "Create or import a data table on a slide, style it, and export "
            "it back out as CSV/JSON."
        ),
        "packs": ["tables-charts"],
        "steps": [
            {"tool": "enable_tools",
             "why": "packs=['tables-charts'] turns on the structural table set"},
            {"tool": "create_table",
             "why": "new table with optional initial data; import_table "
                    "builds one from an existing CSV/JSON file instead"},
            {"tool": "set_table_cells",
             "why": "bulk cell edits: text plus per-cell bold/size/fill in "
                    "one call"},
            {"tool": "merge_cells", "optional": True,
             "why": "span header cells; unmerge_cells reverses it"},
            {"tool": "apply_table_style", "optional": True,
             "why": "one of the 74 built-in PowerPoint table styles by name"},
            {"tool": "export_table", "optional": True,
             "why": "round-trip the finished table to CSV/JSON for records"},
        ],
        "notes": [
            "create_chart / update_chart_data live in the same pack for "
            "bar, column, line, and pie charts.",
            "Cell addresses are 0-based (row, col) into the full grid; merge "
            "continuations refuse edits and name the origin cell.",
        ],
    },
    "template-deck-setup": {
        "summary": (
            "Start a new deck FROM a template (theme, layouts, and masters "
            "intact) and shape its skeleton."
        ),
        "packs": ["design"],
        "steps": [
            {"tool": "enable_tools",
             "why": "packs=['design'] turns on create_presentation and "
                    "slide-size/hide controls"},
            {"tool": "create_presentation",
             "why": "byte-copies the template so theme colors, fonts, "
                    "layouts, and masters all carry over"},
            {"tool": "list_elements",
             "why": "kind='layouts' shows which layouts the template offers, "
                    "by name and part"},
            {"tool": "insert_slide",
             "why": "add slides by layout name or index; repeat per section"},
            {"tool": "set_placeholder_text",
             "why": "titles and body placeholders inherit the template's "
                    "styling automatically"},
            {"tool": "move_slide", "optional": True,
             "why": "adjust running order; set_slide_hidden parks optional "
                    "slides without deleting them"},
        ],
        "notes": [
            "keep_slides=False (default) strips the template's slides and "
            "keeps only its design machinery.",
        ],
    },
    "render-and-review": {
        "summary": (
            "The verification loop: render slides to images, inspect, fix, "
            "re-render; finish with a full-deck validation and PDF."
        ),
        "packs": ["assembly-export", "design"],
        "steps": [
            {"tool": "enable_tools",
             "why": "packs=['assembly-export','design'] turns on export, "
                    "validation, and the autofit report"},
            {"tool": "get_export_engines",
             "why": "confirm which render engine exists here (PowerPoint COM "
                    "is ground truth; LibreOffice drifts on themes/fonts)"},
            {"tool": "export_slide_image",
             "why": "PNG per slide; read the image to catch overflow, "
                    "overlap, and color mistakes"},
            {"tool": "get_autofit_state", "optional": True,
             "why": "flags text boxes that are shrinking or overflowing "
                    "without opening PowerPoint"},
            {"tool": "validate",
             "why": "package payload check plus a real invisible-PowerPoint "
                    "open; repair prompts count as failure"},
            {"tool": "export_pdf",
             "why": "the shareable artifact once the deck passes review"},
        ],
        "notes": [
            "set_notes (same pack) writes the speaker notes reviewers read "
            "alongside the rendered slides.",
        ],
    },
    "batch-edit-from-view": {
        "summary": (
            "The token-cheap editing loop, all in lite: read the anchored "
            "view once, then apply many edits in one atomic call."
        ),
        "packs": ["lite"],
        "steps": [
            {"tool": "get_presentation_view",
             "why": "one anchored markdown projection of the deck: slide "
                    "anchors [s:id], shape anchors [a:hex], table cells "
                    "t:hex:rNcN"},
            {"tool": "apply_edits",
             "why": "batch of {op, anchor, ...} edits; every anchor is "
                    "resolved BEFORE anything mutates, one backup and one "
                    "save for the whole batch"},
            {"tool": "get_presentation_view", "optional": True,
             "why": "re-read only when anchors went stale (shape deleted or "
                    "re-created); surviving anchors stay valid"},
        ],
        "notes": [
            "A stale or ambiguous anchor refuses the WHOLE batch and lists "
            "every failed op index; nothing is half-applied.",
            "search_and_replace inside apply_edits covers deck-wide text "
            "swaps without any anchors at all.",
        ],
    },
}


def get_workflows(task: str | None = None) -> dict:
    """Serve one workflow by name, or the index of all of them."""
    if task is None:
        return {
            "workflows": {
                name: {"summary": wf["summary"], "packs": wf["packs"]}
                for name, wf in WORKFLOWS.items()
            },
            "note": (
                "call get_workflows(task=<name>) for the full step sequence; "
                "enable_tools lists every pack with token costs"
            ),
        }
    if task not in WORKFLOWS:
        raise PptMcpError(
            f"unknown workflow {task!r}; available: {sorted(WORKFLOWS)}"
        )
    return {"task": task, **WORKFLOWS[task]}
