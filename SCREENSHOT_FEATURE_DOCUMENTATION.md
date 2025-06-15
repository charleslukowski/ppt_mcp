# PowerPoint Screenshot Feature

## Overview

The PowerPoint MCP Server now includes a **screenshot feature** that captures each slide of a PowerPoint presentation as high-quality image files. This feature is particularly useful for AI vision analysis, slide review, and creating visual documentation.

## Platform Requirements

- **Windows Only**: This feature requires Windows operating system
- **Microsoft PowerPoint**: PowerPoint must be installed on the system
- **pywin32**: Python Windows extensions package

## Installation

The Windows dependency is automatically installed when you install the requirements:

```bash
pip install -r requirements.txt
```

For manual installation of Windows dependencies:

```bash
pip install pywin32
```

## Usage

### Basic Usage

```python
# Take screenshots of all slides with default settings
result = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx"
})
```

### Advanced Usage

```python
# Custom screenshot settings
result = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "output_dir": "screenshots/",           # Custom output directory
    "image_format": "PNG",                  # Image format (PNG, JPG, etc.)
    "width": 1920,                          # Screenshot width in pixels
    "height": 1080                          # Screenshot height in pixels
})
```

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | - | Path to the PowerPoint file |
| `output_dir` | string | No | temp directory | Directory to save screenshots |
| `image_format` | string | No | "PNG" | Image format (PNG, JPG, etc.) |
| `width` | integer | No | 1920 | Screenshot width in pixels |
| `height` | integer | No | 1080 | Screenshot height in pixels |

## Output

The tool returns a JSON response containing:

```json
{
  "total_slides": 5,
  "screenshot_paths": [
    "/path/to/slide_001.png",
    "/path/to/slide_002.png",
    "/path/to/slide_003.png",
    "/path/to/slide_004.png",
    "/path/to/slide_005.png"
  ],
  "image_format": "PNG",
  "dimensions": "1920x1080",
  "output_directory": "/path/to/screenshots"
}
```

## Use Cases

### 1. AI Vision Analysis
```python
# Take screenshots for AI analysis
screenshots = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "width": 2048,
    "height": 1536
})

# Use with vision AI models to:
# - Analyze slide layouts
# - Extract text from images
# - Identify charts and diagrams
# - Generate slide summaries
```

### 2. Quality Assurance
```python
# Create thumbnails for presentation review
thumbnails = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "output_dir": "review_thumbnails/",
    "width": 800,
    "height": 600
})
```

### 3. Documentation
```python
# Generate slide images for documentation
docs = await session.call_tool("screenshot_slides", {
    "file_path": "training_presentation.pptx",
    "output_dir": "documentation/slides/",
    "image_format": "JPG"
})
```

## Technical Details

### How It Works

1. **COM Interface**: Uses Windows COM to control PowerPoint application
2. **Slide Export**: Each slide is exported using PowerPoint's native export functionality
3. **High Quality**: Screenshots maintain original slide quality and formatting
4. **Automatic Cleanup**: Temporary files and COM objects are automatically cleaned up

### File Naming

Screenshots are automatically named with the pattern:
- `slide_001.png`
- `slide_002.png` 
- `slide_003.png`
- etc.

### Supported Formats

- **Input**: All PowerPoint formats (.pptx, .ppt, .pptm)
- **Output**: PNG, JPG, BMP, GIF, TIFF

## Error Handling

The feature includes comprehensive error handling for:

- Missing PowerPoint installation
- File access permissions
- COM interface errors
- Invalid file paths
- Unsupported image formats

## Performance Considerations

- **PowerPoint Launch**: PowerPoint application will briefly open during the process
- **Memory Usage**: Large presentations may require significant memory
- **Processing Time**: ~1-2 seconds per slide depending on complexity
- **File Size**: High-resolution screenshots can be large (1-5MB per slide)

## Testing

Use the provided test script to verify functionality:

```bash
python test_screenshot_feature.py
```

## Troubleshooting

### Common Issues

1. **"Screenshot feature is only available on Windows"**
   - Solution: Use Windows system with PowerPoint installed

2. **"win32com not available"**
   - Solution: Install pywin32 package: `pip install pywin32`

3. **"PowerPoint file not found"**
   - Solution: Verify file path exists and is accessible

4. **"COM interface errors"**
   - Solution: Ensure PowerPoint is properly installed and not running

### Debug Tips

- Enable verbose logging to see detailed COM interactions
- Check Windows Event Viewer for PowerPoint-related errors
- Verify file permissions for input and output directories
- Test with simple presentations first

## Security Considerations

- PowerPoint application runs with user privileges
- Screenshot files inherit directory permissions
- Temporary files are automatically cleaned up
- No network access required (local operation only)

## Future Enhancements

Potential future improvements:
- Batch processing multiple presentations
- Custom slide selection (specific slide ranges)
- Metadata extraction (slide notes, comments)
- Integration with cloud storage services
- Automated slide annotation 