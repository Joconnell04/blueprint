#!/bin/bash
# Build a Claude Desktop extension bundle (.mcpb) that installs by drag-and-drop in Settings > Extensions.
set -euo pipefail
cd "$(dirname "$0")/.."
npm run build >/dev/null
rm -rf bundle && mkdir -p bundle
cp -R dist builder bundle/
[ -f assets/icon.png ] || python3 assets/make_icon.py >/dev/null
cp assets/icon.png bundle/icon.png
rm -f bundle/builder/*.xlsx bundle/builder/test_build.py
rm -rf bundle/builder/__pycache__
cat > bundle/package.json <<JSON
{
  "name": "blueprint",
  "version": "$(node -p "require('./package.json').version")",
  "type": "module",
  "private": true,
  "dependencies": $(node -p "JSON.stringify(Object.fromEntries(Object.entries(require('./package.json').dependencies).filter(([k]) => !['express','cors','@modelcontextprotocol/express','@modelcontextprotocol/node'].includes(k))))")
}
JSON
cp package-lock.json bundle/ 2>/dev/null || true
(cd bundle && npm install --omit=dev --ignore-scripts --no-audit --no-fund >/dev/null 2>&1 && rm -f package-lock.json)
cat > bundle/manifest.json <<JSON
{
  "manifest_version": "0.3",
  "name": "blueprint",
  "display_name": "Blueprint",
  "icon": "icon.png",
  "version": "$(node -p "require('./package.json').version")",
  "description": "Plan a workbook or a deck on a canvas with Claude, approve it, then build the real .xlsx or .pptx.",
  "long_description": "Three layers in both modes: Claude proposes the outline (tabs or slides), you edit and approve it on an inline canvas; Claude proposes each layout as blocks on a grid, you drag, resize, annotate and approve; Claude answers every note and builds the file.",
  "author": { "name": "Jackson O'Connell" },
  "server": {
    "type": "node",
    "entry_point": "dist/server/main.js",
    "mcp_config": {
      "command": "node",
      "args": ["\${__dirname}/dist/server/main.js"],
      "env": {
        "BLUEPRINT_OUTPUT": "\${user_config.output_dir}",
        "BLUEPRINT_PYTHON": "\${user_config.python}"
      }
    }
  },
  "user_config": {
    "output_dir": {
      "type": "directory",
      "title": "Output folder",
      "description": "Where built files are saved",
      "default": "\${HOME}/Documents/Blueprint",
      "required": false
    },
    "python": {
      "type": "string",
      "title": "Python interpreter",
      "description": "Python 3 with openpyxl and python-pptx (pip3 install openpyxl python-pptx)",
      "default": "$(python3 -c "import sys; print(sys.executable)")",
      "required": false
    }
  },
  "tools": [
    { "name": "start_project", "description": "Start a workbook or deck plan" },
    { "name": "propose_outline", "description": "Propose the tabs or slides and render the canvas" },
    { "name": "propose_layouts", "description": "Propose the blocks on each tab or slide" },
    { "name": "update_plan", "description": "Apply chat edits to the plan" },
    { "name": "get_plan", "description": "Read the plan with hand edits and open notes" },
    { "name": "resolve_notes", "description": "Answer or decline the user's notes" },
    { "name": "build_workbook", "description": "Build the .xlsx from the approved plan" },
    { "name": "build_deck", "description": "Build the .pptx from the approved plan" },
    { "name": "review_deck", "description": "Slide images of the last build, for a visual check" },
    { "name": "review_workbook", "description": "Sheet images of the last build, for a visual check" },
    { "name": "import_workbook", "description": "Scan an existing .xlsx into a plan" },
    { "name": "import_deck", "description": "Scan an existing .pptx into a plan" },
    { "name": "extract_format", "description": "Make a style from example files" },
    { "name": "restyle_workbook", "description": "Reformat an imported workbook in place" },
    { "name": "list_templates", "description": "List saved plan templates" },
    { "name": "save_template", "description": "Save this plan as a template" },
    { "name": "apply_template", "description": "Start a plan from a template" },
    { "name": "show_plan", "description": "Re-render the canvas" },
    { "name": "list_projects", "description": "List saved plans" }
  ],
  "keywords": ["excel", "xlsx", "workbook", "spreadsheet", "powerpoint", "pptx", "deck", "slides", "planning", "mcp-app"],
  "license": "MIT",
  "compatibility": { "claude_desktop": ">=0.10.0", "platforms": ["darwin"], "runtimes": { "node": ">=20.0.0" } }
}
JSON
rm -f blueprint.mcpb
npx -y @anthropic-ai/mcpb@2.1.2 pack bundle blueprint.mcpb
ls -la blueprint.mcpb
