# KitchenSink4PPT

Everything plus the kitchen sink for Microsoft PowerPoint: an MCP server for
.pptx files, engineered not to corrupt. Slides, text, tables, charts, notes,
export, and the one thing no other server in the ecosystem does: arbitrary
vector graphics as native, editable PowerPoint shapes.

## The headline: real graphics, not pictures of graphics

Feed `svg_to_shapes` any SVG and it compiles into a grouped tree of NATIVE
PowerPoint shapes: custom geometry, gradients, strokes, text, nested groups.
No rasterizing, no external APIs, no PowerPoint install needed to build.
Every shape keeps its own id, so afterward you edit the diagram
semantically: recolor node 7, resize the box, rewrite its label.

Connectors are GLUED. Move the node and the arrows follow, exactly as if a
person had drawn the diagram by hand in PowerPoint. The human you hand the
deck to can nudge any piece in seconds, with no regeneration round trip.

`insert_shape` (presets and freeform geometry), `insert_connector`
(straight, elbow, curved, arrowheads), `group_shapes`, `align_shapes`,
`distribute_shapes`, and `set_z_order` cover from-scratch diagram building;
`export_slide_image` renders a PNG so the agent can look at what it made and
fix it.

## Install

```
pip install kitchensink4ppt
```

MCP config (Claude Desktop, Claude Code, or any MCP client):

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "ppt-mcp"
    }
  }
}
```

Requires Python 3.12+. Everything file-based runs on any OS; PDF and image
export prefer PowerPoint via COM on Windows and fall back to LibreOffice
headless where available. Nothing ever needs a network connection.

## Tiered loading: start light, grow mid-session

The server starts in lite mode: 20 tools, roughly 3.4k tokens of tool
context, covering reading, slide CRUD, text, batch editing, backups, and
diagnostics. The other 47 tools are registered but disabled until asked for:

```
enable_tools(packs=["graphics"])
```

The tool list grows in place (the server emits tools/list_changed and the
client re-fetches), and the result reports the approximate token cost added
so the tradeoff is visible. `disable_tools` shrinks the surface again, and
`get_workflows` ships recipes that name the right pack for each job.

Environment pins for hosts and power users:

| Variable | Effect |
|---|---|
| `KS4P_MODE` | startup surface: `lite` (default), `full`, or a pack list like `graphics,com` |
| `KS4P_PACK_POLICY` | `auto` (default) or `locked` (enable_tools refuses; surface fixed at startup) |
| `KS4P_ALLOWED_ROOTS` | opt-in path sandbox; tools refuse to touch files outside these roots |

## Pack inventory (67 tools total)

| Pack | Tools | ~Tokens | What is in it |
|---|---|---|---|
| lite core (always on) | 20 | 3.4k | anchored deck view, atomic batch edits, get/find/replace text, slide insert/delete/duplicate/reorder, placeholder text, info and enumeration, copy, snapshots, backups, diagnose, workflows, enable/disable_tools |
| graphics | 13 | 3.3k | shapes, glued connectors, SVG compiler, groups, align/distribute, z-order, text boxes, run formatting, bullets |
| tables-charts | 17 | 3.4k | create table, bulk cells, merge/unmerge, row and column insert/delete, borders and fills, widths/heights, 74 built-in styles, CSV/JSON export/import, bar/line/pie charts with editable data workbooks |
| design | 5 | 0.8k | create presentation FROM template, slide size, hide slide, move slide, autofit overflow report |
| assembly-export | 10 | 1.4k | speaker notes, footers and slide numbers, PDF export, per-slide PNG render, engine detection, opens-clean validation, full text extraction |
| com | 2 | 0.2k | PowerPoint status and zombie process check (Windows) |

Full surface: about 12.6k tokens if you pin `KS4P_MODE=full`.

Structural table operations on the file itself (merging, inserting and
deleting rows AND columns, per-edge borders) exist in no other PowerPoint
MCP server; they were previously COM-or-nothing.

## Safety story

The same discipline as KitchenSink4Word, applied from day one:

- **Atomic validated saves.** Every mutation rebuilds the package in
  memory, validates the payload, then atomically replaces the file. A
  failed operation leaves the original byte-identical; the file is never
  absent from its own path, even for an instant.
- **Two-slot backups.** Before each mutation the current content rotates
  into `prev.pptx` and `anchor.pptx` under a hidden `.ks4p-backups/`
  folder. `manage_backups` lists, restores (undoably), and purges;
  `create_snapshot` makes DTG-stamped permanent keepers.
- **Byte-identical passthrough.** Parts the tool did not touch are written
  back byte-for-byte, so themes, media, and animations survive edits to
  other slides untouched.
- **Conservative refusals.** Ambiguous targets refuse with a candidate
  list (no first-match guessing), merged-cell surgery that would split a
  span refuses, stale batch anchors refuse the whole batch before anything
  mutates. Refusals are structured (`{ok: false, error: {code, message,
  hint}}`) and tell you the exact next call to make.
- **Sandboxing, opt-in.** Set `KS4P_ALLOWED_ROOTS` and every path in and
  out is checked.
- **Locks.** Mutations of one file are serialized in-process and across
  processes; files open in PowerPoint are refused rather than corrupted.

## Maturity

Beta. The file layer (packages, slides, text, graphics, tables, charts,
notes, export) is covered by 320+ tests, including validation that
generated decks open clean in real PowerPoint, and the server passes a raw
stdio protocol round-trip suite. It has not yet had a long field life;
treat important decks with the respect the backup tools make easy, and
expect fast point releases. Live editing of a deck open in PowerPoint,
comments, and transitions/animations are on the roadmap, not in v1.

## License

AGPL-3.0-only (see LICENSE and NOTICE.md). Free for personal, academic,
and open-source use. If you want to embed it in closed-source commercial
software, open an issue to discuss a commercial license.

## Family

Sibling of [KitchenSink4Word](https://github.com/nometalalchemist/KitchenSink4Word)
(the same engineering for .docx; [site](https://nometalalchemist.github.io/KitchenSink4Word/)).
