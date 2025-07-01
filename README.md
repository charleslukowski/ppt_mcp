# PowerPoint MCP Server

A comprehensive Model Context Protocol (MCP) server for PowerPoint presentation manipulation using python-pptx. This server provides programmatic access to PowerPoint operations through the MCP protocol, optimized for use with AI assistants like Claude in Cursor IDE.

## Features

### Core Presentation Management
- **`create_presentation`** - Create presentations from scratch or templates
- **`load_presentation`** - Load existing PowerPoint files for editing
- **`save_presentation`** - Save presentations with automatic path resolution
- **`add_slide`** - Add slides with different layouts (title, content, blank, etc.)
- **`extract_text`** - Extract all text content for analysis
- **`get_presentation_info`** - Get metadata and structure information

### Content Creation & Manipulation
- **`add_text_box`** - Rich text formatting (font size, bold, italic, colors)
- **`add_image`** - Images from files or URLs with precise positioning
- **`add_chart`** - Data-driven charts (column, bar, line, pie, area)
- **`create_from_json`** - Schema-driven presentation creation
- **`add_professional_shape`** - Professional shapes from built-in library

### Visual Analysis & Review
- **`screenshot_slides`** - Generate high-quality slide screenshots (Windows only)
- **`critique_presentation`** - Comprehensive analysis (design, content, accessibility, technical)

### Style Management & Automation
- **`analyze_presentation_style`** - Extract style patterns for learning
- **`create_style_profile`** - Build reusable style profiles
- **`apply_style_profile`** - Apply learned styles to presentations
- **`save_style_profile`** / **`load_style_profile`** - Persist style data

### Professional Layout & Design
- **`create_layout_grid`** - Professional alignment grids
- **`snap_to_grid`** - Snap shapes to grid positions
- **`distribute_shapes`** - Even spacing distribution
- **`create_color_palette`** / **`apply_color_palette`** - Brand-consistent colors
- **`create_typography_profile`** / **`apply_typography_style`** - Typography hierarchies

### Master Themes & Templates
- **`create_master_slide_theme`** - Master slide themes with consistent formatting
- **`apply_master_theme`** - Apply themes across entire presentations
- **`create_template`** - Reusable templates with placeholders and logic
- **`apply_template`** - Data-driven template application
- **`bulk_generate_presentations`** - Generate multiple presentations from templates


## Installation

1. **Install dependencies:**
```bash
pip install -r requirements.txt
```

2. **Verify python-pptx installation:**
```bash
python -c "import pptx; print('python-pptx installed successfully')"
```

## Usage

### Starting the Server

```bash
python powerpoint_mcp_server.py
```

The server will start and listen for MCP protocol messages via stdio.

### Integration with Cursor

1. Copy `cursor_config.json.example` to `cursor_config.json`
2. Update the path in the configuration if needed
3. Add to your Cursor MCP settings

Example configuration:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["./powerpoint_mcp_server.py"],
      "env": {},
      "description": "PowerPoint MCP Server for presentation manipulation"
    }
  }
}
```


## Architecture

### MCP Integration
- **Resources**: Exposes presentations as readable resources
- **Tools**: Provides comprehensive PowerPoint manipulation tools
- **Protocol**: Full MCP compliance with proper error handling

## Limitations

- **Memory usage**: Presentations are kept in memory during operations
- **File formats**: Only supports .pptx format (PowerPoint 2007+)
- **Concurrent access**: Single-threaded operation
- **Template complexity**: Advanced template features may not be fully supported

## License

This project is licensed under the MIT License.

## Acknowledgments

- Built on the [python-pptx](https://python-pptx.readthedocs.io/) library

