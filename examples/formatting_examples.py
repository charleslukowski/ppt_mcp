#!/usr/bin/env python3
"""
PowerPoint MCP Server - Formatting Examples

This script demonstrates the comprehensive formatting capabilities added to the
PowerPoint MCP server, including:

- Text formatting (fonts, colors, alignment, styles)
- Shape formatting (borders, fills)
- Background formatting
- Existing text modification

Run this script to see all formatting features in action.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from powerpoint_mcp_server_stable import StablePowerPointManager

def create_title_slide():
    """Create a title slide with advanced formatting"""
    print("📝 Creating title slide with advanced formatting...")
    
    manager = StablePowerPointManager()
    prs_id = manager.create_presentation()
    
    # Add title slide with custom background
    slide_index = manager.add_slide(prs_id, 6)  # Blank layout
    manager.set_slide_background(prs_id, slide_index, background_color="#1E3A8A")
    
    # Main title - large, bold, centered
    manager.add_text_box(
        prs_id, slide_index,
        text="Advanced Formatting Demo",
        left=1, top=2, width=8, height=2,
        font_size=48, font_name="Arial", font_color="white",
        bold=True, text_alignment="center"
    )
    
    # Subtitle - smaller, italic, centered
    manager.add_text_box(
        prs_id, slide_index,
        text="Showcasing PowerPoint MCP Server Formatting Features",
        left=1, top=4.5, width=8, height=1,
        font_size=20, font_name="Calibri", font_color="#E5E7EB",
        italic=True, text_alignment="center"
    )
    
    return manager, prs_id

def create_content_slide(manager, prs_id):
    """Create a content slide demonstrating various text formatting"""
    print("📝 Creating content slide with various text formatting...")
    
    slide_index = manager.add_slide(prs_id, 6)
    manager.set_slide_background(prs_id, slide_index, background_color="#F8FAFC")
    
    # Slide title
    manager.add_text_box(
        prs_id, slide_index,
        text="Text Formatting Options",
        left=1, top=0.5, width=8, height=1,
        font_size=36, font_name="Arial", font_color="#1F2937",
        bold=True, text_alignment="center"
    )
    
    # Feature examples with different alignments and colors
    features = [
        {"text": "Left-aligned bold text in red", "align": "left", "color": "#DC2626", "bold": True},
        {"text": "Center-aligned italic text in blue", "align": "center", "color": "#2563EB", "italic": True},
        {"text": "Right-aligned underlined text in green", "align": "right", "color": "#059669", "underline": True},
        {"text": "Justified text with multiple formatting options applied simultaneously", "align": "justify", "color": "#7C3AED", "bold": True, "italic": True}
    ]
    
    for i, feature in enumerate(features):
        y_pos = 2 + (i * 1.2)
        manager.add_text_box(
            prs_id, slide_index,
            text=feature["text"],
            left=1, top=y_pos, width=8, height=1,
            font_size=18, font_name="Calibri", font_color=feature["color"],
            bold=feature.get("bold", False),
            italic=feature.get("italic", False),
            underline=feature.get("underline", False),
            text_alignment=feature["align"]
        )

def create_styled_boxes_slide(manager, prs_id):
    """Create a slide with styled text boxes (borders and fills)"""
    print("📝 Creating styled boxes slide...")
    
    slide_index = manager.add_slide(prs_id, 6)
    
    # Slide title
    manager.add_text_box(
        prs_id, slide_index,
        text="Styled Text Boxes",
        left=1, top=0.5, width=8, height=1,
        font_size=36, font_name="Arial", font_color="#1F2937",
        bold=True, text_alignment="center"
    )
    
    # Box 1: Info box with blue theme
    manager.add_text_box(
        prs_id, slide_index,
        text="Information Box\nWith blue background and dark border",
        left=0.5, top=2, width=3.5, height=2,
        font_size=16, font_color="white", text_alignment="center",
        fill_color="#3B82F6", border_color="#1E40AF", border_width=3
    )
    
    # Box 2: Warning box with yellow theme
    manager.add_text_box(
        prs_id, slide_index,
        text="Warning Box\nWith yellow background and orange border",
        left=4.5, top=2, width=3.5, height=2,
        font_size=16, font_color="#92400E", text_alignment="center",
        fill_color="#FEF3C7", border_color="#F59E0B", border_width=2
    )
    
    # Box 3: Success box with green theme
    manager.add_text_box(
        prs_id, slide_index,
        text="Success Box\nWith green background and matching border",
        left=2, top=4.5, width=4, height=2,
        font_size=16, font_color="white", text_alignment="center",
        fill_color="#10B981", border_color="#047857", border_width=2
    )

def create_font_showcase_slide(manager, prs_id):
    """Create a slide showcasing different fonts"""
    print("📝 Creating font showcase slide...")
    
    slide_index = manager.add_slide(prs_id, 6)
    manager.set_slide_background(prs_id, slide_index, background_color="#F3F4F6")
    
    # Slide title
    manager.add_text_box(
        prs_id, slide_index,
        text="Font Showcase",
        left=1, top=0.5, width=8, height=1,
        font_size=36, font_name="Arial", font_color="#111827",
        bold=True, text_alignment="center"
    )
    
    # Different fonts with their names
    fonts = [
        {"name": "Arial", "color": "#EF4444"},
        {"name": "Calibri", "color": "#10B981"},
        {"name": "Times New Roman", "color": "#3B82F6"},
        {"name": "Georgia", "color": "#8B5CF6"},
        {"name": "Trebuchet MS", "color": "#F59E0B"}
    ]
    
    for i, font in enumerate(fonts):
        y_pos = 2 + (i * 1)
        manager.add_text_box(
            prs_id, slide_index,
            text=f"This text is rendered in {font['name']} font",
            left=1, top=y_pos, width=8, height=0.8,
            font_size=20, font_name=font["name"], font_color=font["color"],
            text_alignment="left"
        )

def create_before_after_slide(manager, prs_id):
    """Create a slide demonstrating text formatting modification"""
    print("📝 Creating before/after formatting slide...")
    
    slide_index = manager.add_slide(prs_id, 6)
    
    # Slide title
    manager.add_text_box(
        prs_id, slide_index,
        text="Before & After Text Formatting",
        left=1, top=0.5, width=8, height=1,
        font_size=32, font_name="Arial", font_color="#1F2937",
        bold=True, text_alignment="center"
    )
    
    # Before text (simple formatting)
    manager.add_text_box(
        prs_id, slide_index,
        text="BEFORE: Plain text with basic formatting",
        left=1, top=2, width=8, height=1,
        font_size=18, font_name="Arial", font_color="#6B7280",
        text_alignment="left"
    )
    
    # After text (enhanced formatting)
    manager.add_text_box(
        prs_id, slide_index,
        text="AFTER: Enhanced text with colors, borders, and styling",
        left=1, top=4, width=8, height=1.5,
        font_size=20, font_name="Calibri", font_color="white",
        bold=True, italic=True, text_alignment="center",
        fill_color="#8B5CF6", border_color="#6D28D9", border_width=3
    )
    
    # Demo the format_existing_text function
    # Get the first text box and modify it
    content = manager.list_slide_content(prs_id, slide_index)
    if content["shape_count"] >= 2:
        # Modify the "BEFORE" text to show the difference
        manager.format_existing_text(
            prs_id, slide_index, 1,  # Second shape (BEFORE text)
            font_size=16, font_color="#DC2626", italic=True
        )

def main():
    """Create a complete presentation showcasing all formatting features"""
    print("🚀 Creating comprehensive formatting demonstration...")
    print("=" * 60)
    
    # Create presentation with multiple slides
    manager, prs_id = create_title_slide()
    create_content_slide(manager, prs_id)
    create_styled_boxes_slide(manager, prs_id)
    create_font_showcase_slide(manager, prs_id)
    create_before_after_slide(manager, prs_id)
    
    # Save the presentation
    output_file = "formatting_showcase.pptx"
    saved_path = manager.save_presentation(prs_id, output_file)
    
    print("=" * 60)
    print("✅ Formatting showcase presentation created!")
    print(f"📁 Saved as: {saved_path}")
    print("\n🎨 Features demonstrated:")
    print("  • Text formatting: fonts, sizes, colors, styles")
    print("  • Text alignment: left, center, right, justify")
    print("  • Text boxes: borders, fills, backgrounds")
    print("  • Slide backgrounds: colors")
    print("  • Dynamic formatting modification")
    print("\n📝 Open the file in PowerPoint to see all formatting features!")

if __name__ == "__main__":
    main() 