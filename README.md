# KitchenSink4PPT

Everything plus the kitchen sink for Microsoft PowerPoint: an MCP server
for .pptx files, engineered not to corrupt.

Status: pre-release scaffold (0.1.0.dev0). The safety core is in place:
atomic validated saves, two-slot backups, opt-in path sandboxing, and a
package layer that writes untouched parts back byte-for-byte identical.
The tool surface arrives in subsequent releases.

License: AGPL-3.0-only (see LICENSE and NOTICE.md).
