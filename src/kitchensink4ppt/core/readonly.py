"""Which tools can change something, and which cannot.

Every registered tool declares a readOnlyHint annotation, and this module
is where that decision lives. Claude Desktop groups tools by the hint:
the read-only ones land in a "Read-only tools" group the user can approve
once, and everything else keeps asking before it runs. That makes the
classification a safety boundary rather than documentation, so it is
declared here by name and never inferred from what a tool is called.

The bar for true is that the tool cannot change ANYTHING: not the
deck, not any other file (temporary files included), not a process,
not what this server exposes. One exclusion is worth naming, because it
reads as a read: validate asks PowerPoint whether the deck opens clean,
which starts a hidden PowerPoint and opens the file in it. The live route
the dual-mode readers take is different: it attaches to a PowerPoint that
is already running and reads, so those tools stay read-only. The export
tools all write a file, so none of them qualify however read-shaped they
feel.

The sibling web server (kitchensink4web) set the pattern and the
discipline: an honest hint buys client-side permission lenience, and an
optimistic one would be a false safety claim in metadata.
"""

from __future__ import annotations


#: Tools that cannot change anything: no deck, no file, no process,
#: no server state. These carry readOnlyHint: true.
READ_ONLY: frozenset[str] = frozenset({
    "audit_accessibility", "check_layout", "comment_report",
    "compare_decks", "deck_statistics", "diagnose", "extract_brand",
    "extract_text", "find_text", "font_inventory", "get_autofit_state",
    "get_chart_data", "get_document_properties", "get_export_engines",
    "get_footer_support", "get_media_playback", "get_notes",
    "get_presentation_info",
    "get_presentation_view", "get_slide_info", "get_table", "get_text",
    "get_theme", "get_transitions", "get_workflows", "list_animations",
    "list_comments", "list_elements", "list_equations",
    "list_hyperlinks", "list_master_elements", "live_status",
    "powerpoint_status", "zombie_check",
})

#: Everything else, declared explicitly rather than inferred. A tool is
#: here because it writes a deck, writes a file, starts or drives an
#: Office process, or changes what this server exposes.
MUTATING: frozenset[str] = frozenset({
    "add_comment", "add_entrance_animation", "add_equation_to_shape",
    "add_layout_placeholder", "align_shapes", "anonymize_deck",
    "apply_brand", "apply_edits", "apply_layout", "apply_table_style",
    "clear_animations", "compress_deck", "copy_format", "copy_position",
    "copy_presentation", "copy_slide_between", "create_chart",
    "create_layout", "create_presentation", "create_snapshot",
    "create_table", "delete_comment", "delete_master_shape",
    "delete_notes", "delete_shape", "delete_slide", "delete_table_cols",
    "delete_table_rows", "disable_tools", "distribute_shapes",
    "duplicate_slide", "enable_tools", "export_handout", "export_pdf",
    "export_slide_image", "export_table", "fit_text", "format_chart",
    "format_table_cells", "format_text", "generate_agenda_slide",
    "generate_diagram", "group_shapes", "import_table", "insert_audio",
    "insert_connector", "insert_equation", "insert_image",
    "insert_master_shape", "insert_shape", "insert_slide",
    "insert_table_cols", "insert_table_rows", "insert_textbox",
    "insert_video", "live_save", "live_scroll_to", "manage_backups",
    "manage_custom_show", "manage_section", "merge_cells",
    "merge_decks", "move_slide", "refresh_agenda_slide",
    "remove_hyperlink", "remove_layout_placeholder", "reorder_slides",
    "replace_colors", "replace_fonts", "replace_image",
    "replace_image_everywhere", "reply_to_comment", "resolve_comment",
    "search_and_replace", "set_alt_text", "set_bullets",
    "set_column_widths", "set_diagram_text", "set_document_properties",
    "set_footer", "set_hyperlink", "set_image", "set_language",
    "set_layout_placeholder", "set_master_background",
    "set_master_placeholder", "set_media_playback", "set_notes",
    "set_placeholder_text",
    "set_reading_order", "set_row_heights", "set_shape",
    "set_show_properties", "set_slide_background", "set_slide_hidden",
    "set_slide_size", "set_table_cells", "set_theme_colors",
    "set_theme_fonts", "set_transition", "set_z_order", "split_deck",
    "svg_to_shapes", "ungroup_shapes", "unmerge_cells",
    "update_chart_data", "validate",
})


def read_only_hint(name: str) -> bool:
    """The MCP readOnlyHint for one tool.

    Unknown names RAISE at registration time rather than defaulting.
    Defaulting to false would quietly drop a read tool out of the
    client's read-only group; defaulting to true would put a tool that
    can change a file into a group the user bulk-approves. Neither is a
    decision this module is willing to make on someone's behalf."""
    if name in READ_ONLY:
        return True
    if name in MUTATING:
        return False
    raise RuntimeError(
        f"tool {name!r} is not classified in core/readonly.py. Every "
        f"tool must be declared READ_ONLY or MUTATING before it can be "
        f"registered, because the annotation drives a bulk-approval "
        f"group in the client and an unclassified tool would land in it "
        f"by accident."
    )
