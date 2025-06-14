# PowerPoint MCP Server

A comprehensive Model Context Protocol (MCP) server for PowerPoint presentation manipulation using python-pptx. This server provides programmatic access to PowerPoint operations through the MCP protocol, optimized for use with AI assistants like Claude in Cursor IDE.

## Features

### Core Presentation Management
- **Create presentations** from scratch or templates
- **Load existing presentations** from files
- **Save presentations** to various formats
- **Extract text content** from presentations
- **Get presentation metadata** and structure information

### Slide Operations
- **Add slides** with different layouts (title, content, blank, etc.)
- **Manage slide content** programmatically
- **Batch slide creation** from structured data

### Content Manipulation
- **Text boxes** with rich formatting (font size, bold, italic)
- **Images** from files or URLs with precise positioning
- **Charts** (column, bar, line, pie) with data-driven content
- **Precise positioning** using inch-based measurements

### Advanced Features
- **Template-based workflows** for consistent branding
- **JSON schema-driven** presentation creation
- **Batch processing** capabilities
- **URL image integration** for dynamic content
- **Content extraction** for analysis and processing

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

## Available Tools

### 1. create_presentation
Create a new PowerPoint presentation, optionally from a template.

```json
{
  "name": "create_presentation",
  "arguments": {
    "template_path": "optional/path/to/template.pptx"
  }
}
```

### 2. load_presentation
Load an existing PowerPoint presentation from file.

```json
{
  "name": "load_presentation",
  "arguments": {
    "file_path": "path/to/presentation.pptx"
  }
}
```

### 3. save_presentation
Save a presentation to file.

```json
{
  "name": "save_presentation",
  "arguments": {
    "presentation_id": "prs_0",
    "file_path": "output/presentation.pptx"
  }
}
```

### 4. add_slide
Add a new slide to a presentation.

```json
{
  "name": "add_slide",
  "arguments": {
    "presentation_id": "prs_0",
    "layout_index": 6
  }
}
```

**Layout indices:**
- 0: Title slide
- 1: Title and content
- 2: Section header
- 3: Two content
- 4: Comparison
- 5: Title only
- 6: Blank
- 7: Content with caption
- 8: Picture with caption

### 5. add_text_box
Add a formatted text box to a slide.

```json
{
  "name": "add_text_box",
  "arguments": {
    "presentation_id": "prs_0",
    "slide_index": 0,
    "text": "Hello, World!",
    "left": 1,
    "top": 1,
    "width": 8,
    "height": 1,
    "font_size": 24,
    "bold": true,
    "italic": false
  }
}
```

### 6. add_image
Add an image from file or URL.

```json
{
  "name": "add_image",
  "arguments": {
    "presentation_id": "prs_0",
    "slide_index": 0,
    "image_source": "https://example.com/image.jpg",
    "left": 2,
    "top": 2,
    "width": 4,
    "height": 3
  }
}
```

### 7. add_chart
Add a data-driven chart to a slide.

```json
{
  "name": "add_chart",
  "arguments": {
    "presentation_id": "prs_0",
    "slide_index": 0,
    "chart_type": "column",
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series_data": {
      "Sales": [100, 150, 120, 180],
      "Profit": [20, 30, 25, 40]
    }
  }
}
```

### 8. extract_text
Extract all text content from a presentation.

```json
{
  "name": "extract_text",
  "arguments": {
    "presentation_id": "prs_0"
  }
}
```

### 9. get_presentation_info
Get metadata about a presentation.

```json
{
  "name": "get_presentation_info",
  "arguments": {
    "presentation_id": "prs_0"
  }
}
```

### 10. create_from_json
Create a presentation from structured JSON data.

```json
{
  "name": "create_from_json",
  "arguments": {
    "json_data": {
      "slide1": {
        "title": "Welcome",
        "content": "This is the first slide"
      },
      "slide2": {
        "title": "Data Overview",
        "content": "Key metrics and insights"
      }
    },
    "template_path": "optional/template.pptx"
  }
}
```

## Examples

### Example 1: Create a Simple Presentation

```python
# Through MCP calls:
# 1. Create presentation
create_presentation()  # Returns: prs_0

# 2. Add title slide
add_slide(presentation_id="prs_0", layout_index=0)  # Returns: slide 0

# 3. Add title text
add_text_box(
    presentation_id="prs_0",
    slide_index=0,
    text="My Presentation",
    left=1, top=1, width=8, height=1,
    font_size=32, bold=True
)

# 4. Save presentation
save_presentation(presentation_id="prs_0", file_path="my_presentation.pptx")
```

### Example 2: Data-Driven Presentation

```python
# Create presentation with chart
create_presentation()  # Returns: prs_0
add_slide(presentation_id="prs_0", layout_index=6)  # Blank slide

# Add chart with sales data
add_chart(
    presentation_id="prs_0",
    slide_index=0,
    chart_type="column",
    categories=["Jan", "Feb", "Mar", "Apr"],
    series_data={
        "Revenue": [10000, 12000, 11000, 15000],
        "Expenses": [8000, 9000, 8500, 11000]
    }
)

save_presentation(presentation_id="prs_0", file_path="sales_report.pptx")
```

### Example 3: Template-Based Workflow

```python
# Create from template
create_presentation(template_path="company_template.pptx")  # Returns: prs_0

# Add content slides
add_slide(presentation_id="prs_0", layout_index=1)  # Title and content
add_text_box(
    presentation_id="prs_0",
    slide_index=0,
    text="Q4 Results",
    font_size=28, bold=True
)

# Add image from URL
add_image(
    presentation_id="prs_0",
    slide_index=0,
    image_source="https://charts.example.com/q4-performance.png",
    left=1, top=3, width=8, height=4
)
```

## Resources

The server exposes loaded presentations as MCP resources:

- **URI**: `powerpoint://prs_0`
- **Type**: `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- **Content**: JSON representation of presentation structure and text content

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

- Built on the excellent [python-pptx](https://python-pptx.readthedocs.io/) library
- Follows MCP protocol specifications
