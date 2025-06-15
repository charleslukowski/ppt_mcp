# PowerPoint MCP Server - Prompt System Review

## Overview

I've added a comprehensive prompt system to your PowerPoint MCP server that provides contextual guidance for professional presentation creation. This system includes:

1. **Overall server system prompt**
2. **Category-specific prompts** for different types of operations
3. **Tool-specific prompts** for individual functions
4. **Contextual guidance** that adapts based on parameters
5. **Logging integration** to track guidance usage

## System Architecture

### Main System Prompt

```
You are a professional PowerPoint presentation assistant with expertise in creating 
visually appealing, well-structured, and content-rich presentations. Your role is to:

1. PROFESSIONAL DESIGN: Always prioritize clean, professional layouts with consistent 
   formatting, appropriate spacing, and visual hierarchy.

2. CONTENT CLARITY: Focus on clear, concise messaging with effective use of bullet points, 
   headings, and visual elements to enhance understanding.

3. BRAND CONSISTENCY: Maintain consistent colors, fonts, and styling throughout presentations.
   Use corporate color palettes and professional typography.

4. ACCESSIBILITY: Ensure presentations are readable with appropriate contrast, font sizes,
   and logical content flow.

5. EFFICIENCY: Leverage templates, themes, and automation to create presentations quickly
   while maintaining high quality standards.

Key Principles:
- Less is more: Avoid cluttered slides
- Visual hierarchy: Use size, color, and positioning to guide attention
- Consistency: Maintain uniform styling across all slides
- Professional aesthetics: Choose appropriate colors, fonts, and layouts
- Data visualization: Present complex information through charts and graphics
```

## Category-Specific Prompts

### 1. Presentation Creation
```
When creating presentations:
- Start with a clear purpose and target audience in mind
- Use appropriate slide layouts for different content types
- Establish consistent branding and visual themes early
- Consider the presentation flow and logical progression
- Ensure all slides serve the overall narrative
```

### 2. Content Formatting
```
When formatting content:
- Use bullet points for lists and key points
- Keep text concise and readable (max 6 bullet points per slide)
- Use appropriate font sizes (title: 28-36pt, body: 18-24pt)
- Maintain consistent spacing and alignment
- Use bold/italic strategically for emphasis
- Avoid full sentences in bullet points when possible
```

### 3. Visual Design
```
When designing visual elements:
- Use high-contrast color combinations for readability
- Align elements to invisible grids for professional appearance
- Maintain consistent spacing between elements
- Use white space effectively to avoid clutter
- Choose colors that support the message and brand
- Ensure images are high-quality and appropriately sized
```

### 4. Charts & Data Visualization
```
When creating charts and data visualizations:
- Choose chart types that best represent the data story
- Use clear, descriptive titles and labels
- Apply consistent color schemes across all charts
- Avoid 3D effects that can distort data perception
- Include data source attribution when appropriate
- Keep axis labels readable and well-formatted
```

### 5. Template & Automation
```
When working with templates and automation:
- Design templates with flexibility and reusability in mind
- Use clear placeholder naming conventions
- Include conditional logic for dynamic content
- Maintain consistent formatting across generated slides
- Test templates with various data sets before deployment
- Document template usage and data requirements
```

## Tool-Specific Prompts

### Key Tools with Custom Prompts:

1. **create_presentation**: Guidance for professional setup and template selection
2. **add_text_box**: Typography and content hierarchy guidance
3. **add_image**: Image quality and integration guidance
4. **add_chart**: Data visualization best practices
5. **create_color_palette**: Brand consistency and accessibility
6. **apply_typography_style**: Content hierarchy and readability
7. **create_master_slide_theme**: Comprehensive design consistency
8. **create_template**: Reusable design structures
9. **bulk_generate_presentations**: Efficiency with quality maintenance

## Contextual Guidance System

The system provides dynamic guidance based on operation parameters:

### Text Box Guidance
- **Large fonts (≥28pt)**: "This appears to be title text - use strong, impactful language."
- **Small fonts (≤14pt)**: "This appears to be caption text - keep it concise and supportive."

### Chart Guidance
- **Pie charts**: "Pie charts work best with 5 or fewer categories. Consider using a bar chart for more categories."
- **Line charts**: "Line charts are excellent for showing trends over time. Ensure data points are clearly labeled."

### Color Palette Guidance
- **General**: "Consider the presentation's purpose: corporate (blues/grays), creative (varied), financial (blues/greens)."

## Implementation Details

### New Methods Added:

1. `get_system_prompt()` - Returns the main system prompt
2. `get_category_prompt(category)` - Returns category-specific guidance
3. `get_tool_prompt(tool_name)` - Returns tool-specific guidance  
4. `get_contextual_guidance(tool_name, **kwargs)` - Returns contextual guidance with parameters
5. `log_operation_guidance(tool_name, **kwargs)` - Logs guidance for each operation

### Integration Points:

The following methods now include prompt guidance logging:
- `create_presentation()` - Logs presentation creation guidance
- `add_text_box()` - Logs text formatting guidance with font size context
- `add_chart()` - Logs chart creation guidance with type and category count
- `add_image()` - Logs image addition guidance with URL detection
- `create_color_palette()` - Logs color palette guidance with customization info

## Usage Examples

### Accessing Prompts Programmatically:

```python
# Get the main system prompt
system_prompt = manager.get_system_prompt()

# Get category-specific guidance
formatting_guidance = manager.get_category_prompt("formatting")

# Get tool-specific guidance
chart_guidance = manager.get_tool_prompt("add_chart")

# Get contextual guidance with parameters
text_guidance = manager.get_contextual_guidance(
    "add_text_box", 
    font_size=32, 
    text_length=50
)
```

### Automatic Logging:

When operations are performed, guidance is automatically logged:
```
INFO - Operation guidance for add_text_box: Add text with appropriate formatting...
INFO - Operation guidance for add_chart: Create data visualizations that clearly communicate...
```

## Benefits

1. **Consistency**: Ensures all presentations follow professional standards
2. **Education**: Teaches best practices through contextual guidance
3. **Quality**: Improves presentation quality through automated suggestions
4. **Efficiency**: Reduces decision-making time with clear guidance
5. **Adaptability**: Provides context-aware suggestions based on parameters

## Next Steps

Consider adding:
1. **Dynamic prompts** based on presentation type (business, academic, creative)
2. **Industry-specific guidance** (finance, healthcare, technology)
3. **Accessibility prompts** for inclusive design
4. **Collaborative prompts** for team presentations
5. **Brand-specific prompts** loaded from configuration files

## Review Questions

1. Do the prompts align with your presentation style and requirements?
2. Are there specific industry or use-case prompts you'd like to add?
3. Should any prompts be more or less prescriptive?
4. Would you like additional contextual guidance for specific scenarios?
5. Should the logging be more or less verbose?

---

**Note**: All prompts are now integrated into the PowerPoint MCP server and will provide guidance whenever the corresponding tools are used. 