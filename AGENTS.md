# multiCAD-mcp Agent Guidelines

Guide for AI Coding Assistants (Antigravity IDE, Claude, Cursor, Copilot, Cline, etc.) interacting with or developing `multiCAD-mcp`.

---

## 1. Project Architecture

`multiCAD-mcp` is a Model Context Protocol (MCP) server built with **FastMCP** that interfaces with Windows CAD applications (AutoCAD, ZWCAD, GstarCAD, BricsCAD) via **Windows COM** (`win32com`).

### Unified Tools (7 tools replacing 56 legacy commands)

| Tool | Purpose | Primary Format |
|------|---------|----------------|
| `manage_session` | Connection lifecycle, viewport navigation, undo/redo, dashboard | `operations` (JSON/shorthand) or `action` param |
| `draw_entities` | Geometric entities (lines, circles, rects, arcs, polylines, text, tables, leaders) | Shorthand string or list of dicts |
| `manage_layers` | Create, list, delete, rename, visibility toggle, color | Shorthand string or list of dicts |
| `manage_files` | Save (DWG/DXF/PDF), new, close, switch drawings | Shorthand string or list of dicts |
| `manage_entities` | Select, move, rotate, scale, color, copy, paste, delete | Shorthand string or list of dicts |
| `manage_blocks` | Create blocks, insert, list, query attributes | Shorthand string or list of dicts |
| `export_data` | Extract entity metadata (JSON) or export to Excel | `scope="all"|"selected"`, `format="json"|"excel"` |

---

## 2. Tool Input Formats

All tools accept **flexible input formats** to prevent validation errors:

### A. Shorthand Format (Recommended for speed and minimal tokens)
```text
# draw_entities
rect|0,0,0|420,297,0|cyan
text|10,10|Title Block|5|cyan
line|0,0|100,100|red|walls
circle|50,50|25|blue

# manage_files
save|floorplan.dwg
save|export.dxf|dxf

# manage_layers
create|walls|red|50
on|walls,furniture
off|defpoints

# manage_entities
select|layer|walls
move|A1B2,C3D4|10|5
rotate|A1B2|45|0|0
```

### B. Direct Keyword Arguments (for `manage_session`)
```python
# Direct action
manage_session(action="status")
manage_session(action="connect")
manage_session(action="zoom_extents")
manage_session(action="undo", count=2)

# Or batch operations list
manage_session(operations=[{"action": "zoom_extents"}])
```

### C. JSON Objects or Lists
Any tool can receive standard Python/JSON lists of objects or single dictionaries:
```python
draw_entities(entities=[
    {"type": "rectangle", "corner1": "0,0", "corner2": "420,297", "color": "cyan"}
])
```

---

## 3. Windows COM & CAD Environments

- **OS Requirement**: Windows only (`win32com`, `pythoncom`).
- **Connection Strategy**:
  - The server attempts `GetActiveObject(prog_id)` up to 3 times to attach to a running CAD GUI.
  - If not active in ROT, it falls back to `Dispatch(prog_id)` and sets `app.Visible = True`.
- **Supported CAD ProgIDs**:
  - `autocad`: `AutoCAD.Application`
  - `zwcad`: `ZWCAD.Application`
  - `gcad`: `GCAD.Application`
  - `bricscad`: `BricscadApp.AcadApplication`
- **Important Gotcha**: In Windows COM, if CAD processes are created in the background, ensure you don't accumulate orphan processes (`ZWCAD.exe` / `acad.exe` with `/Automation -Embedding`). Clean up when done via `manage_session(action="disconnect")`.

---

## 4. Development & Testing Workflow

- **Package Manager**: Use `uv`.
  ```powershell
  # Sync all dependencies including dev tools (pytest, ruff, mypy)
  uv sync --dev

  # Run test suite
  uv run pytest tests/ -v

  # Code formatting & linting
  uv run ruff format src/
  uv run ruff check src/
  ```
- **Do not run `python -m pip install` inside uv virtualenvs**: Virtualenvs created by `uv` do not have `pip` by default. Use `uv pip install` or `uv add` if extra packages are needed.

---

## 5. Client Integrations

- **Antigravity IDE**: Configuration at `.agents/mcp_config.json` and `%USERPROFILE%\.gemini\config\mcp_config.json`.
- **VS Code**: Configuration at `.vscode/mcp.json`.
- **Claude Desktop**: Configuration at `%APPDATA%\Claude\claude_desktop_config.json`.
- **Entry point**:
  `command`: `<project_path>\.venv\Scripts\python.exe`
  `args`: `["<project_path>\\src\\server.py"]`
