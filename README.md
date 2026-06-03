# Toolsopedia

Toolsopedia is a collection of small web tools, each living in its own folder at the project root.

## Structure

Each tool is a self-contained app directory under `tools/` (for example, `tools/pdf-to-excel-app/`).
As new tools are added, they should be placed in their own folders alongside existing tools under `tools/`.

## Current Tools

- `tools/pdf-to-excel-app/`: Flask app that converts PDFs to Excel files.

## Conventions

- Keep each tool isolated (its own dependencies, templates, static assets, etc.).
- Shared assets or utilities can be added later under a dedicated shared folder if needed.
