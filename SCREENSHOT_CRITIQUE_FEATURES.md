# PowerPoint MCP Server - Screenshot & Critique Features

## Overview

The PowerPoint MCP Server now includes advanced **screenshot** and **self-critique** functionality, enabling comprehensive presentation analysis and visual review capabilities. These features work together to provide AI-powered quality assessment and visual documentation of PowerPoint presentations.

## 🖼️ Screenshot Feature

### Description
Generates high-quality screenshots of all slides in a PowerPoint presentation using Windows COM automation. Perfect for AI vision analysis, presentation review, and documentation.

### Requirements
- **Windows Operating System** (COM automation required)
- **Microsoft PowerPoint** installed
- **pywin32** Python package

### Usage

```python
# Basic screenshot generation
result = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx"
})

# Advanced configuration
result = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "output_dir": "screenshots/",
    "image_format": "PNG",
    "width": 1920,
    "height": 1080
})
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | ✅ | - | Path to PowerPoint file |
| `output_dir` | string | ❌ | temp dir | Screenshot output directory |
| `image_format` | string | ❌ | "PNG" | Image format (PNG, JPG, etc.) |
| `width` | integer | ❌ | 1920 | Screenshot width in pixels |
| `height` | integer | ❌ | 1080 | Screenshot height in pixels |

### Output
```json
{
  "total_slides": 5,
  "screenshot_paths": [
    "/path/to/slide_001.png",
    "/path/to/slide_002.png",
    "..."
  ],
  "image_format": "PNG",
  "dimensions": "1920x1080",
  "output_directory": "/path/to/screenshots"
}
```

## 🔍 Critique Feature

### Description
Comprehensive AI-powered analysis that evaluates presentations across multiple quality dimensions, providing scores, identifying issues, and suggesting improvements.

### Analysis Categories

#### 1. **Design Analysis**
- Font consistency and usage
- Font size appropriateness (18pt+ recommended)
- Visual element detection
- Layout quality assessment
- Color usage patterns

#### 2. **Content Analysis**
- Text quantity per slide (300 chars max recommended)
- Slide structure and titles
- Bullet point counts (5-7 max recommended)
- Empty slide detection
- Content organization

#### 3. **Accessibility Analysis**
- Alt text for images
- Color contrast assessment
- Text readability
- Visual impairment considerations

#### 4. **Technical Analysis**
- File size optimization
- Slide count management
- Embedded object analysis
- Performance considerations

### Usage

```python
# Comprehensive critique (all categories)
result = await session.call_tool("critique_presentation", {
    "file_path": "presentation.pptx",
    "critique_type": "comprehensive",
    "include_screenshots": true
})

# Specific category analysis
result = await session.call_tool("critique_presentation", {
    "file_path": "presentation.pptx",
    "critique_type": "design",
    "include_screenshots": false
})
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | ✅ | - | Path to PowerPoint file |
| `critique_type` | string | ❌ | "comprehensive" | Analysis type |
| `include_screenshots` | boolean | ❌ | true | Generate screenshots |
| `output_dir` | string | ❌ | temp dir | Screenshot directory |

### Critique Types
- `"design"` - Font, layout, visual consistency
- `"content"` - Text, structure, organization  
- `"accessibility"` - Alt text, contrast, readability
- `"technical"` - File size, performance, optimization
- `"comprehensive"` - All categories combined

### Output Structure

```json
{
  "file_path": "presentation.pptx",
  "critique_type": "comprehensive",
  "timestamp": "2024-01-15T10:30:00",
  "summary": {
    "total_slides": 10,
    "overall_score": 75.2,
    "assessment": "Good",
    "critical_issues": 2,
    "warnings": 5,
    "recommendations": 8,
    "analysis_categories": ["design", "content", "accessibility", "technical"]
  },
  "issues": [
    {
      "type": "warning",
      "category": "design",
      "slide": 3,
      "issue": "Small font sizes detected",
      "description": "Minimum font size is 12pt. Consider 18pt+ for readability."
    }
  ],
  "strengths": [
    "Consistent font usage throughout presentation",
    "Good overall design consistency"
  ],
  "recommendations": [
    "Review font consistency across slides",
    "Ensure minimum 18pt font size for readability"
  ],
  "detailed_analysis": {
    "design": {
      "score": 78,
      "metrics": {
        "total_fonts": 2,
        "font_sizes_range": {"min": 12, "max": 36, "avg": 20.5}
      }
    }
  },
  "screenshots": ["/path/to/slide_001.png", "..."]
}
```

## 🔗 Integrated Workflow

The screenshot and critique features work seamlessly together:

1. **Visual Analysis**: Screenshots enable AI vision models to analyze slide layouts, charts, and visual elements
2. **Comprehensive Assessment**: Critique combines structural analysis with visual inspection
3. **Rich Output**: Results include both analytical data and visual references
4. **Quality Assurance**: Perfect for automated presentation review workflows

### Example Workflow

```python
# Step 1: Generate presentation screenshots
screenshots = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "output_dir": "analysis/"
})

# Step 2: Run comprehensive critique with screenshots
critique = await session.call_tool("critique_presentation", {
    "file_path": "presentation.pptx",
    "critique_type": "comprehensive",
    "include_screenshots": true,
    "output_dir": "analysis/"
})

# Step 3: Review results
print(f"Overall Score: {critique['summary']['overall_score']}/100")
print(f"Assessment: {critique['summary']['assessment']}")
print(f"Screenshots: {len(critique['screenshots'])} generated")
```

## 📊 Scoring System

### Score Ranges
- **90-100**: Excellent - Professional quality, minimal issues
- **80-89**: Good - Solid presentation with minor improvements needed
- **70-79**: Fair - Acceptable quality with some issues to address
- **60-69**: Needs Improvement - Multiple issues requiring attention
- **<60**: Poor - Significant issues requiring major revision

### Issue Types
- **🔴 Critical**: Major problems that significantly impact quality
- **⚠️ Warning**: Issues that should be addressed for best practices
- **💡 Recommendation**: Suggestions for enhancement

## 🎯 Use Cases

### 1. **AI Vision Analysis**
```python
# Generate screenshots for AI model analysis
screenshots = await session.call_tool("screenshot_slides", {
    "file_path": "presentation.pptx",
    "width": 2048,
    "height": 1536
})
# Feed screenshots to vision AI for layout analysis, text extraction, etc.
```

### 2. **Quality Assurance Pipeline**
```python
# Automated presentation review
critique = await session.call_tool("critique_presentation", {
    "file_path": "presentation.pptx",
    "critique_type": "comprehensive"
})
if critique["summary"]["overall_score"] < 70:
    print("Presentation needs improvement before approval")
```

### 3. **Batch Analysis**
```python
# Analyze multiple presentations
presentations = ["pres1.pptx", "pres2.pptx", "pres3.pptx"]
for pres in presentations:
    critique = await session.call_tool("critique_presentation", {
        "file_path": pres,
        "critique_type": "comprehensive"
    })
    print(f"{pres}: {critique['summary']['overall_score']}/100")
```

### 4. **Documentation Generation**
```python
# Create visual documentation
screenshots = await session.call_tool("screenshot_slides", {
    "file_path": "training_deck.pptx",
    "output_dir": "docs/images/",
    "image_format": "JPG"
})
# Use screenshots in documentation, wikis, etc.
```

## ⚡ Performance

- **Screenshot Generation**: ~1-2 seconds per slide
- **Critique Analysis**: ~2-5 seconds for comprehensive analysis
- **Memory Usage**: Minimal impact, automatic cleanup
- **File Size**: Screenshots typically 1-5MB each (PNG format)

## 🛠️ Technical Details

### Screenshot Implementation
- Uses Windows COM (`win32com.client`) to control PowerPoint
- Exports slides using PowerPoint's native high-quality export
- Automatic file naming (`slide_001.png`, `slide_002.png`, etc.)
- Supports all PowerPoint formats (.pptx, .ppt, .pptm)

### Critique Implementation
- Analyzes presentation using `python-pptx` library
- Multi-dimensional scoring algorithm
- Configurable analysis depth
- JSON-structured results for easy integration

### Error Handling
- Comprehensive error catching and reporting
- Graceful degradation when features unavailable
- Automatic resource cleanup
- Platform compatibility checks

## 🔧 Configuration

### Screenshot Quality Settings
```python
# High-resolution for detailed analysis
{"width": 2560, "height": 1440, "image_format": "PNG"}

# Optimized for web/preview
{"width": 1280, "height": 720, "image_format": "JPG"}

# Thumbnail generation
{"width": 640, "height": 360, "image_format": "PNG"}
```

### Critique Sensitivity Tuning
The critique system uses configurable thresholds:
- Font size minimum: 18pt (adjustable)
- Text length maximum: 300 characters per slide
- Bullet point maximum: 7 per slide
- File size warning: 50MB

## 🔮 Future Enhancements

### Planned Features
- **Batch Processing**: Multiple presentations at once
- **Custom Scoring Rules**: User-defined quality criteria
- **Visual AI Integration**: Advanced image analysis
- **Report Generation**: PDF/HTML critique reports
- **Template Analysis**: Best practice template suggestions

### Potential Integrations
- Vision AI models for layout analysis
- OCR for text extraction from images
- Accessibility scanning tools
- Brand compliance checking
- A/B testing capabilities

## 🤝 Contributing

The screenshot and critique features are designed to be extensible:

1. **Custom Analysis Rules**: Add new critique categories in `_analyze_*` methods
2. **Output Formats**: Extend screenshot formats in `screenshot_slides`
3. **Scoring Algorithms**: Modify scoring logic in `_calculate_critique_summary`
4. **Integration Hooks**: Add callbacks for external AI services

## 📝 Examples

See `test_screenshot_critique.py` for comprehensive usage examples and testing patterns.

---

*These features transform the PowerPoint MCP Server into a comprehensive presentation analysis platform, enabling AI-powered quality assurance and visual documentation workflows.* 