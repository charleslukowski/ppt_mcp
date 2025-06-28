#!/usr/bin/env python3
"""
PowerPoint MCP Server - Stable Production Version

This version includes core improvements from our development but maintains
stability and compatibility with Cursor's settings UI.

Key improvements included:
- Input validation (basic, no Pydantic dependency)
- Enhanced success messages
- Core 5 essential tools
- Post-processing fixes
- Simple error handling
"""

import asyncio
import json
import logging
import os
import sys
import tempfile
from typing import Any, Dict, List, Optional

# Configure logging
logging.basicConfig(level=logging.INFO, stream=sys.stderr)
logger = logging.getLogger("powerpoint-mcp-stable")

# Core dependencies only
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
except ImportError as e:
    print(f"python-pptx library not found: {e}")
    print("Please install with: pip install python-pptx")
    sys.exit(1)

try:
    from mcp.server import Server
    from mcp.server.models import InitializationOptions
    from mcp.server.stdio import stdio_server
    from mcp.types import Tool, TextContent, EmbeddedResource
except ImportError as e:
    print(f"MCP library not found: {e}")
    print("Please install with: pip install mcp")
    sys.exit(1)

# =============================================================================
# SIMPLIFIED INPUT VALIDATION (No Pydantic dependency)
# =============================================================================

def validate_basic_args(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
    """Basic input validation without external dependencies"""
    
    # Get presentation_id if present
    prs_id = arguments.get("presentation_id")
    if prs_id and not isinstance(prs_id, str):
        raise ValueError("presentation_id must be a string")
    
    # Get slide_index if present
    slide_index = arguments.get("slide_index")
    if slide_index is not None:
        if not isinstance(slide_index, int) or slide_index < 0:
            raise ValueError("slide_index must be a non-negative integer")
    
    # Tool-specific validation
    if tool_name == "add_text_box":
        text = arguments.get("text", "")
        if not text or not isinstance(text, str):
            raise ValueError("text must be a non-empty string")
        
        font_size = arguments.get("font_size", 18)
        if not isinstance(font_size, int) or font_size < 8 or font_size > 72:
            raise ValueError("font_size must be between 8 and 72")
        
        font_name = arguments.get("font_name", "Calibri")
        if not isinstance(font_name, str):
            raise ValueError("font_name must be a string")
        
        text_alignment = arguments.get("text_alignment", "left")
        valid_alignments = ["left", "center", "right", "justify"]
        if text_alignment.lower() not in valid_alignments:
            raise ValueError(f"text_alignment must be one of: {valid_alignments}")
        
        border_width = arguments.get("border_width", 0)
        if not isinstance(border_width, (int, float)) or border_width < 0:
            raise ValueError("border_width must be a non-negative number")
    
    elif tool_name == "add_image":
        image_source = arguments.get("image_source", "")
        if not image_source or not isinstance(image_source, str):
            raise ValueError("image_source must be a non-empty string")
    
    elif tool_name == "add_chart":
        chart_type = arguments.get("chart_type", "")
        valid_types = ["column", "bar", "line", "pie", "area"]
        if chart_type not in valid_types:
            raise ValueError(f"chart_type must be one of: {valid_types}")
        
        categories = arguments.get("categories", [])
        if not categories or not isinstance(categories, list):
            raise ValueError("categories must be a non-empty list")
        
        series_data = arguments.get("series_data", {})
        if not series_data or not isinstance(series_data, dict):
            raise ValueError("series_data must be a non-empty dictionary")
    
    elif tool_name == "save_presentation":
        file_path = arguments.get("file_path", "")
        if not file_path or not isinstance(file_path, str):
            raise ValueError("file_path must be a non-empty string")
    
    elif tool_name == "load_presentation":
        file_path = arguments.get("file_path", "")
        if not file_path or not isinstance(file_path, str):
            raise ValueError("file_path must be a non-empty string")
    
    elif tool_name == "add_slide":
        layout_index = arguments.get("layout_index", 6)
        if not isinstance(layout_index, int) or layout_index < 0:
            raise ValueError("layout_index must be a non-negative integer")
    
    elif tool_name == "extract_text":
        # No additional validation needed beyond presentation_id
        pass
    
    elif tool_name == "get_presentation_info":
        # No additional validation needed beyond presentation_id
        pass
    
    elif tool_name == "delete_shape":
        shape_index = arguments.get("shape_index")
        if shape_index is None or not isinstance(shape_index, int) or shape_index < 0:
            raise ValueError("shape_index must be a non-negative integer")
    
    elif tool_name == "delete_slide":
        # slide_index validation already handled above
        pass
    
    elif tool_name == "clear_slide":
        # slide_index validation already handled above  
        pass
    
    elif tool_name == "list_slide_content":
        # slide_index validation already handled above
        pass
    
    elif tool_name == "format_existing_text":
        shape_index = arguments.get("shape_index")
        if shape_index is None or not isinstance(shape_index, int) or shape_index < 0:
            raise ValueError("shape_index must be a non-negative integer")
        
        # Validate formatting parameters if provided
        font_size = arguments.get("font_size")
        if font_size is not None and (not isinstance(font_size, int) or font_size < 8 or font_size > 72):
            raise ValueError("font_size must be between 8 and 72")
        
        text_alignment = arguments.get("text_alignment")
        if text_alignment is not None:
            valid_alignments = ["left", "center", "right", "justify"]
            if text_alignment.lower() not in valid_alignments:
                raise ValueError(f"text_alignment must be one of: {valid_alignments}")
    
    elif tool_name == "set_slide_background":
        background_color = arguments.get("background_color")
        background_image = arguments.get("background_image")
        if not background_color and not background_image:
            raise ValueError("Either background_color or background_image must be provided")
    
    return arguments

# =============================================================================
# ENHANCED SUCCESS MESSAGES
# =============================================================================

def format_success_message(tool_name: str, **kwargs) -> str:
    """Generate specific, actionable success messages"""
    
    if tool_name == "add_text_box":
        slide_idx = kwargs.get('slide_index', 0)
        font_size = kwargs.get('font_size', 18)
        font_name = kwargs.get('font_name', 'Calibri')
        text_alignment = kwargs.get('text_alignment', 'left')
        font_color = kwargs.get('font_color')
        fill_color = kwargs.get('fill_color')
        text_preview = kwargs.get('text', '')[:40] + ('...' if len(kwargs.get('text', '')) > 40 else '')
        
        # Build formatting description
        format_desc = f"{font_size}pt {font_name}, {text_alignment} aligned"
        if font_color:
            format_desc += f", color: {font_color}"
        if fill_color:
            format_desc += f", background: {fill_color}"
        
        return f"✅ Added formatted text box to slide {slide_idx + 1}: \"{text_preview}\" ({format_desc})"
    
    elif tool_name == "add_image":
        slide_idx = kwargs.get('slide_index', 0)
        image_source = kwargs.get('image_source', '')
        image_name = os.path.basename(image_source) if image_source else 'image'
        return f"✅ Added image to slide {slide_idx + 1}: {image_name}"
    
    elif tool_name == "add_chart":
        slide_idx = kwargs.get('slide_index', 0)
        chart_type = kwargs.get('chart_type', 'chart')
        categories = kwargs.get('categories', [])
        series_data = kwargs.get('series_data', {})
        return f"✅ Added {chart_type} chart to slide {slide_idx + 1}: {len(categories)} categories, {len(series_data)} series"
    
    elif tool_name == "save_presentation":
        file_path = kwargs.get('file_path', '')
        if file_path:
            # Show both filename and full path for clarity
            file_name = os.path.basename(file_path)
            # Normalize path for display
            display_path = os.path.normpath(file_path)
            
            # Check if it's in Documents folder and mention it prominently
            if "Documents" in display_path:
                return f"✅ Saved presentation: {file_name}\n📁 Location: Documents folder\n📍 Full path: {display_path}"
            else:
                # Truncate very long paths for readability but keep them informative
                if len(display_path) > 80:
                    display_path = f"...{display_path[-77:]}"
                return f"✅ Saved presentation: {file_name}\n📁 Full path: {display_path}"
        return f"✅ Saved presentation → Ready for use!"
    
    elif tool_name == "create_presentation":
        prs_id = kwargs.get('presentation_id', 'new')
        return f"✅ Created presentation {prs_id} → Ready to add slides!"
    
    elif tool_name == "load_presentation":
        prs_id = kwargs.get('presentation_id', 'loaded')
        file_path = kwargs.get('file_path', '')
        file_name = os.path.basename(file_path) if file_path else 'presentation'
        slide_count = kwargs.get('slide_count', 'unknown')
        return f"📂 Loaded presentation {prs_id} from {file_name} → {slide_count} slides available"
    
    elif tool_name == "add_slide":
        slide_idx = kwargs.get('slide_index', 0)
        layout_idx = kwargs.get('layout_index', 6)
        layout_name = kwargs.get('layout_name', f'Layout {layout_idx}')
        total_slides = kwargs.get('total_slides', 'unknown')
        return f"➕ Added slide {slide_idx + 1} using {layout_name} → {total_slides} slides total"
    
    elif tool_name == "extract_text":
        slide_count = kwargs.get('slide_count', 0)
        text_items = kwargs.get('text_items', 0)
        return f"📝 Extracted text from {slide_count} slides → Found {text_items} text items"
    
    elif tool_name == "get_presentation_info":
        slide_count = kwargs.get('slide_count', 0)
        total_shapes = kwargs.get('total_shapes', 0)
        return f"ℹ️ Presentation info: {slide_count} slides, {total_shapes} total shapes"
    
    elif tool_name == "delete_shape":
        slide_idx = kwargs.get('slide_index', 0)
        shape_idx = kwargs.get('shape_index', 0)
        shape_type = kwargs.get('shape_type', 'shape')
        return f"🗑️ Deleted {shape_type} (index {shape_idx}) from slide {slide_idx + 1}"
    
    elif tool_name == "delete_slide":
        slide_idx = kwargs.get('slide_index', 0)
        remaining_slides = kwargs.get('remaining_slides', 'unknown')
        return f"🗑️ Deleted slide {slide_idx + 1} → {remaining_slides} slides remaining"
    
    elif tool_name == "clear_slide":
        slide_idx = kwargs.get('slide_index', 0)
        shapes_cleared = kwargs.get('shapes_cleared', 0)
        return f"🧹 Cleared slide {slide_idx + 1} → Removed {shapes_cleared} shapes"
    
    elif tool_name == "list_slide_content":
        slide_idx = kwargs.get('slide_index', 0)
        shape_count = kwargs.get('shape_count', 0)
        return f"📋 Slide {slide_idx + 1} contents: {shape_count} shapes found"
    
    elif tool_name == "format_existing_text":
        slide_idx = kwargs.get('slide_index', 0)
        shape_idx = kwargs.get('shape_index', 0)
        formatted_props = []
        if kwargs.get('font_size'):
            formatted_props.append(f"size: {kwargs['font_size']}pt")
        if kwargs.get('font_name'):
            formatted_props.append(f"font: {kwargs['font_name']}")
        if kwargs.get('font_color'):
            formatted_props.append(f"color: {kwargs['font_color']}")
        if kwargs.get('text_alignment'):
            formatted_props.append(f"align: {kwargs['text_alignment']}")
        props_desc = ", ".join(formatted_props) if formatted_props else "basic formatting"
        return f"🎨 Updated text formatting for shape {shape_idx} on slide {slide_idx + 1}: {props_desc}"
    
    elif tool_name == "set_slide_background":
        slide_idx = kwargs.get('slide_index', 0)
        bg_color = kwargs.get('background_color')
        bg_image = kwargs.get('background_image')
        if bg_color:
            return f"🎨 Set slide {slide_idx + 1} background color: {bg_color}"
        elif bg_image:
            image_name = os.path.basename(bg_image) if bg_image else 'image'
            return f"🎨 Set slide {slide_idx + 1} background image: {image_name}"
        return f"🎨 Updated slide {slide_idx + 1} background"
    
    return f"✅ {tool_name} completed successfully"

# =============================================================================
# SIMPLIFIED POWERPOINT MANAGER
# =============================================================================

class StablePowerPointManager:
    """Simplified PowerPoint manager focused on core functionality"""
    
    def __init__(self):
        self.presentations: Dict[str, Presentation] = {}
        logger.info("PowerPoint manager initialized")
    
    def create_presentation(self) -> str:
        """Create a new blank presentation"""
        prs = Presentation()
        prs_id = f"ppt_{len(self.presentations)}"
        self.presentations[prs_id] = prs
        logger.info(f"Created presentation: {prs_id}")
        return prs_id
    
    def load_presentation(self, file_path: str) -> str:
        """Load an existing PowerPoint presentation from file"""
        # Ensure .pptx extension if not provided
        if not file_path.lower().endswith('.pptx'):
            file_path += '.pptx'
        
        # Convert to absolute path for consistency
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        # Check if file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Presentation file not found: {file_path}")
        
        try:
            # Load the presentation
            prs = Presentation(file_path)
            prs_id = f"ppt_{len(self.presentations)}"
            self.presentations[prs_id] = prs
            
            logger.info(f"Loaded presentation: {prs_id} from {file_path}")
            return prs_id
            
        except Exception as e:
            logger.error(f"Failed to load presentation: {e}")
            raise RuntimeError(f"Failed to load presentation from {file_path}: {e}")
    
    def add_slide(self, prs_id: str, layout_index: int = 6) -> int:
        """Add a new slide to the presentation with specified layout"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        # Validate layout index
        if layout_index < 0 or layout_index >= len(prs.slide_layouts):
            available_layouts = len(prs.slide_layouts)
            raise ValueError(f"Layout index {layout_index} is invalid. Available layouts: 0-{available_layouts-1}")
        
        # Add the slide
        layout = prs.slide_layouts[layout_index]
        slide = prs.slides.add_slide(layout)
        slide_index = len(prs.slides) - 1
        
        logger.info(f"Added slide {slide_index} with layout {layout_index} to {prs_id}")
        return slide_index
    
    def add_text_box(self, prs_id: str, slide_index: int, text: str, 
                     left: float = 1, top: float = 1, width: float = 8, height: float = 1,
                     font_size: int = 18, font_name: str = "Calibri", font_color: Optional[str] = None,
                     bold: bool = False, italic: bool = False, underline: bool = False,
                     text_alignment: str = "left", fill_color: Optional[str] = None,
                     border_color: Optional[str] = None, border_width: float = 0) -> bool:
        """Add a text box to a slide with comprehensive formatting"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        # Add slide if needed
        while len(prs.slides) <= slide_index:
            layout = prs.slide_layouts[6]  # Blank layout
            prs.slides.add_slide(layout)
        
        slide = prs.slides[slide_index]
        
        # Add text box
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        
        text_frame = textbox.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        
        # Map text alignment
        alignment_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY
        }
        paragraph_alignment = alignment_map.get(text_alignment.lower(), PP_ALIGN.LEFT)
        
        # Apply text formatting
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = paragraph_alignment
            for run in paragraph.runs:
                # Font properties
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.font.bold = bold
                run.font.italic = italic
                run.font.underline = underline
                
                # Font color
                if font_color:
                    try:
                        rgb = self._parse_color(font_color)
                        run.font.color.rgb = RGBColor(*rgb)
                    except Exception as e:
                        logger.warning(f"Invalid font color '{font_color}': {e}")
        
        # Apply shape formatting
        try:
            # Fill color
            if fill_color:
                try:
                    rgb = self._parse_color(fill_color)
                    textbox.fill.solid()
                    textbox.fill.fore_color.rgb = RGBColor(*rgb)
                except Exception as e:
                    logger.warning(f"Invalid fill color '{fill_color}': {e}")
            
            # Border formatting
            if border_width > 0:
                textbox.line.width = Pt(border_width)
                if border_color:
                    try:
                        rgb = self._parse_color(border_color)
                        textbox.line.color.rgb = RGBColor(*rgb)
                    except Exception as e:
                        logger.warning(f"Invalid border color '{border_color}': {e}")
            else:
                # No border
                textbox.line.fill.background()
                
        except Exception as e:
            logger.warning(f"Failed to apply shape formatting: {e}")
        
        logger.info(f"Added formatted text box to slide {slide_index}")
        return True
    
    def _parse_color(self, color_str: str) -> tuple:
        """Parse color string to RGB tuple. Supports hex (#RRGGBB) and RGB (r,g,b) formats"""
        color_str = color_str.strip()
        
        # Hex format: #RRGGBB or RRGGBB
        if color_str.startswith('#'):
            color_str = color_str[1:]
        
        if len(color_str) == 6:
            try:
                r = int(color_str[0:2], 16)
                g = int(color_str[2:4], 16) 
                b = int(color_str[4:6], 16)
                return (r, g, b)
            except ValueError:
                pass
        
        # RGB format: "r,g,b" or "(r,g,b)"
        if ',' in color_str:
            color_str = color_str.strip('()')
            try:
                parts = [int(x.strip()) for x in color_str.split(',')]
                if len(parts) == 3 and all(0 <= x <= 255 for x in parts):
                    return tuple(parts)
            except ValueError:
                pass
        
        # Predefined colors
        color_map = {
            'black': (0, 0, 0),
            'white': (255, 255, 255),
            'red': (255, 0, 0),
            'darkred': (139, 0, 0),
            'green': (0, 128, 0),
            'darkgreen': (0, 100, 0),
            'blue': (0, 0, 255),
            'darkblue': (0, 0, 139),
            'yellow': (255, 255, 0),
            'orange': (255, 165, 0),
            'purple': (128, 0, 128),
            'gray': (128, 128, 128),
            'grey': (128, 128, 128),
            'lightgray': (211, 211, 211),
            'lightgrey': (211, 211, 211),
            'darkgray': (64, 64, 64),
            'darkgrey': (64, 64, 64)
        }
        
        if color_str.lower() in color_map:
            return color_map[color_str.lower()]
        
        raise ValueError(f"Invalid color format: {color_str}. Use hex (#RRGGBB), RGB (r,g,b), or predefined color names")
    
    def format_existing_text(self, prs_id: str, slide_index: int, shape_index: int,
                           font_size: Optional[int] = None, font_name: Optional[str] = None,
                           font_color: Optional[str] = None, bold: Optional[bool] = None,
                           italic: Optional[bool] = None, underline: Optional[bool] = None,
                           text_alignment: Optional[str] = None) -> bool:
        """Modify formatting of existing text shape"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist")
        
        slide = prs.slides[slide_index]
        
        if shape_index >= len(slide.shapes):
            raise ValueError(f"Shape {shape_index} does not exist")
        
        shape = slide.shapes[shape_index]
        
        # Check if it's a text shape
        if not hasattr(shape, 'text_frame'):
            raise ValueError(f"Shape {shape_index} is not a text shape")
        
        text_frame = shape.text_frame
        
        # Apply text alignment if specified
        if text_alignment:
            alignment_map = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY
            }
            paragraph_alignment = alignment_map.get(text_alignment.lower(), PP_ALIGN.LEFT)
            for paragraph in text_frame.paragraphs:
                paragraph.alignment = paragraph_alignment
        
        # Apply text formatting
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if font_name is not None:
                    run.font.name = font_name
                if font_size is not None:
                    run.font.size = Pt(font_size)
                if bold is not None:
                    run.font.bold = bold
                if italic is not None:
                    run.font.italic = italic
                if underline is not None:
                    run.font.underline = underline
                
                if font_color:
                    try:
                        rgb = self._parse_color(font_color)
                        run.font.color.rgb = RGBColor(*rgb)
                    except Exception as e:
                        logger.warning(f"Invalid font color '{font_color}': {e}")
        
        logger.info(f"Updated formatting for text shape {shape_index} on slide {slide_index}")
        return True
    
    def set_slide_background(self, prs_id: str, slide_index: int, 
                           background_color: Optional[str] = None,
                           background_image: Optional[str] = None) -> bool:
        """Set slide background color or image"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist")
        
        slide = prs.slides[slide_index]
        
        try:
            if background_color:
                # Set background color
                background = slide.background
                fill = background.fill
                fill.solid()
                rgb = self._parse_color(background_color)
                fill.fore_color.rgb = RGBColor(*rgb)
                logger.info(f"Set slide {slide_index} background color to {background_color}")
            
            if background_image:
                # Set background image
                if background_image.startswith(('http://', 'https://')):
                    # Download image temporarily
                    import urllib.request
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                        urllib.request.urlretrieve(background_image, tmp_file.name)
                        image_path = tmp_file.name
                else:
                    image_path = background_image
                    if not os.path.exists(image_path):
                        raise FileNotFoundError(f"Background image not found: {image_path}")
                
                # Apply background image
                background = slide.background
                fill = background.fill
                fill.picture(image_path)
                
                # Clean up temporary file if downloaded
                if background_image.startswith(('http://', 'https://')):
                    os.unlink(image_path)
                
                logger.info(f"Set slide {slide_index} background image")
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to set slide background: {e}")
            raise RuntimeError(f"Failed to set slide background: {e}")
    
    def add_image(self, prs_id: str, slide_index: int, image_source: str,
                  left: float = 1, top: float = 1, width: Optional[float] = None, 
                  height: Optional[float] = None) -> bool:
        """Add an image to a slide"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        # Add slide if needed
        while len(prs.slides) <= slide_index:
            layout = prs.slide_layouts[6]
            prs.slides.add_slide(layout)
        
        slide = prs.slides[slide_index]
        
        try:
            # Handle different image sources
            if image_source.startswith(('http://', 'https://')):
                # Download image temporarily
                import urllib.request
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                    urllib.request.urlretrieve(image_source, tmp_file.name)
                    image_path = tmp_file.name
            else:
                # Local file
                image_path = image_source
                if not os.path.exists(image_path):
                    raise FileNotFoundError(f"Image file not found: {image_path}")
            
            # Add image to slide
            if width and height:
                picture = slide.shapes.add_picture(
                    image_path, Inches(left), Inches(top), Inches(width), Inches(height)
                )
            else:
                picture = slide.shapes.add_picture(
                    image_path, Inches(left), Inches(top)
                )
            
            # Clean up temporary file if downloaded
            if image_source.startswith(('http://', 'https://')):
                os.unlink(image_path)
            
            logger.info(f"Added image to slide {slide_index}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to add image: {e}")
            raise
    
    def add_chart(self, prs_id: str, slide_index: int, chart_type: str, 
                  categories: List[str], series_data: Dict[str, List[float]],
                  left: float = 2, top: float = 2, width: float = 6, height: float = 4.5) -> bool:
        """Add a chart to a slide"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        # Add slide if needed
        while len(prs.slides) <= slide_index:
            layout = prs.slide_layouts[6]
            prs.slides.add_slide(layout)
        
        slide = prs.slides[slide_index]
        
        try:
            # Create chart data
            chart_data = CategoryChartData()
            chart_data.categories = categories
            
            for series_name, values in series_data.items():
                if len(values) != len(categories):
                    raise ValueError(f"Series '{series_name}' has {len(values)} values but {len(categories)} categories")
                chart_data.add_series(series_name, values)
            
            # Map chart type
            chart_type_map = {
                "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "line": XL_CHART_TYPE.LINE,
                "pie": XL_CHART_TYPE.PIE,
                "area": XL_CHART_TYPE.AREA
            }
            
            xl_chart_type = chart_type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
            
            # Add chart to slide
            chart = slide.shapes.add_chart(
                xl_chart_type, Inches(left), Inches(top), 
                Inches(width), Inches(height), chart_data
            ).chart
            
            # Basic chart formatting
            chart.has_legend = True
            chart.legend.position = 2  # Right
            
            logger.info(f"Added {chart_type} chart to slide {slide_index}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to add chart: {e}")
            raise
    
    def save_presentation(self, prs_id: str, file_path: str) -> str:
        """Save presentation to file and return file info - handles Windows paths properly"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        # Ensure .pptx extension first (before path processing)
        if not file_path.lower().endswith('.pptx'):
            file_path += '.pptx'
        
        # Handle relative paths intelligently
        if not os.path.isabs(file_path):
            try:
                cwd = os.getcwd()
                logger.info(f"Current working directory: {cwd}")
                
                # Check if we're in a system/application directory
                problematic_paths = ['AppData', 'cursor', 'Program Files', 'Windows']
                is_system_dir = any(path_part in cwd for path_part in problematic_paths)
                
                if is_system_dir:
                    # Use user's Documents folder for better accessibility
                    documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
                    file_path = os.path.join(documents_dir, file_path)
                    logger.info(f"Using Documents directory instead of system directory: {documents_dir}")
                else:
                    # Use the current working directory
                    file_path = os.path.join(cwd, file_path)
                    logger.info(f"Using current working directory: {cwd}")
                    
            except Exception as e:
                # Fallback to Documents folder
                logger.warning(f"Error determining working directory: {e}")
                documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
                file_path = os.path.join(documents_dir, file_path)
                logger.info(f"Fallback to Documents directory: {documents_dir}")
        
        # Normalize path for Windows (handles both / and \ separators)
        file_path = os.path.normpath(file_path)
        
        # Create directory if needed (always try to create parent directory)
        dir_path = os.path.dirname(file_path)
        if dir_path and not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path, exist_ok=True)
                logger.info(f"Created directory: {dir_path}")
            except OSError as e:
                raise RuntimeError(f"Failed to create directory {dir_path}: {e}")
        
        # Save presentation
        try:
            self.presentations[prs_id].save(file_path)
            logger.info(f"PowerPoint saved to: {file_path}")
        except Exception as e:
            logger.error(f"Save failed: {e}")
            raise RuntimeError(f"Failed to save PowerPoint file to {file_path}: {e}")
        
        # Verify the file was created
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
            logger.info(f"Saved {prs_id} to {file_path} ({file_size} bytes)")
            return file_path
        else:
            raise RuntimeError(f"File was not created at {file_path}")
    
    def delete_shape(self, prs_id: str, slide_index: int, shape_index: int) -> bool:
        """Delete a specific shape from a slide by index"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist (presentation has {len(prs.slides)} slides)")
        
        slide = prs.slides[slide_index]
        
        if shape_index >= len(slide.shapes):
            raise ValueError(f"Shape {shape_index} does not exist (slide has {len(slide.shapes)} shapes)")
        
        # Get shape info for logging before deletion
        shape = slide.shapes[shape_index]
        shape_type = "unknown"
        try:
            if hasattr(shape, 'text_frame'):
                shape_type = "text box"
            elif hasattr(shape, 'chart'):
                shape_type = "chart"
            elif hasattr(shape, 'image'):
                shape_type = "image"
        except:
            pass
        
        # Delete the shape
        shape_element = shape.element
        shape_element.getparent().remove(shape_element)
        
        logger.info(f"Deleted {shape_type} (index {shape_index}) from slide {slide_index}")
        return True
    
    def delete_slide(self, prs_id: str, slide_index: int) -> bool:
        """Delete an entire slide from the presentation"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if len(prs.slides) <= 1:
            raise ValueError("Cannot delete slide - presentation must have at least one slide")
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist (presentation has {len(prs.slides)} slides)")
        
        # Remove the slide
        xml_slides = prs.slides._sldIdLst
        slides = list(xml_slides)
        xml_slides.remove(slides[slide_index])
        
        logger.info(f"Deleted slide {slide_index} from presentation {prs_id}")
        return True
    
    def clear_slide(self, prs_id: str, slide_index: int) -> bool:
        """Clear all content from a slide but keep the slide"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist (presentation has {len(prs.slides)} slides)")
        
        slide = prs.slides[slide_index]
        
        # Count shapes before deletion
        shape_count = len(slide.shapes)
        
        # Delete all shapes (in reverse order to avoid index issues)
        for i in range(len(slide.shapes) - 1, -1, -1):
            try:
                shape = slide.shapes[i]
                shape_element = shape.element
                shape_element.getparent().remove(shape_element)
            except Exception as e:
                logger.warning(f"Could not delete shape {i}: {e}")
        
        logger.info(f"Cleared {shape_count} shapes from slide {slide_index}")
        return True
    
    def list_slide_content(self, prs_id: str, slide_index: int) -> Dict[str, Any]:
        """List all content on a slide for easier deletion targeting"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        if slide_index >= len(prs.slides):
            raise ValueError(f"Slide {slide_index} does not exist (presentation has {len(prs.slides)} slides)")
        
        slide = prs.slides[slide_index]
        content = []
        
        for i, shape in enumerate(slide.shapes):
            shape_info = {
                "index": i,
                "type": "unknown",
                "description": ""
            }
            
            try:
                if hasattr(shape, 'text_frame') and shape.text_frame.text:
                    shape_info["type"] = "text"
                    shape_info["description"] = f"Text: '{shape.text_frame.text[:50]}...'" if len(shape.text_frame.text) > 50 else f"Text: '{shape.text_frame.text}'"
                elif hasattr(shape, 'chart'):
                    shape_info["type"] = "chart"
                    shape_info["description"] = "Chart"
                elif hasattr(shape, 'image'):
                    shape_info["type"] = "image"
                    shape_info["description"] = "Image"
                else:
                    shape_info["type"] = "shape"
                    shape_info["description"] = "Shape"
            except:
                pass
            
            content.append(shape_info)
        
        return {
            "slide_index": slide_index,
            "shape_count": len(content),
            "shapes": content
        }
    
    def extract_text(self, prs_id: str) -> List[Dict[str, Any]]:
        """Extract all text content from the presentation"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        extracted_text = []
        
        for slide_idx, slide in enumerate(prs.slides):
            slide_text = {
                "slide_index": slide_idx,
                "slide_number": slide_idx + 1,
                "text_content": []
            }
            
            for shape_idx, shape in enumerate(slide.shapes):
                try:
                    if hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                        shape_text = {
                            "shape_index": shape_idx,
                            "shape_type": "text",
                            "text": shape.text_frame.text.strip()
                        }
                        slide_text["text_content"].append(shape_text)
                    elif hasattr(shape, 'table'):
                        # Extract text from table cells
                        table_text = []
                        for row in shape.table.rows:
                            row_text = []
                            for cell in row.cells:
                                if cell.text.strip():
                                    row_text.append(cell.text.strip())
                            if row_text:
                                table_text.append(" | ".join(row_text))
                        if table_text:
                            shape_text = {
                                "shape_index": shape_idx,
                                "shape_type": "table",
                                "text": "\n".join(table_text)
                            }
                            slide_text["text_content"].append(shape_text)
                except Exception as e:
                    logger.warning(f"Could not extract text from shape {shape_idx} on slide {slide_idx}: {e}")
            
            extracted_text.append(slide_text)
        
        logger.info(f"Extracted text from {len(extracted_text)} slides in {prs_id}")
        return extracted_text

    def get_presentation_info(self, prs_id: str) -> Dict[str, Any]:
        """Get comprehensive presentation information"""
        if prs_id not in self.presentations:
            raise ValueError(f"Presentation {prs_id} not found")
        
        prs = self.presentations[prs_id]
        
        # Count different types of content
        total_text_boxes = 0
        total_images = 0
        total_charts = 0
        total_shapes = 0
        slide_details = []
        
        for slide_idx, slide in enumerate(prs.slides):
            slide_info = {
                "slide_index": slide_idx,
                "slide_number": slide_idx + 1,
                "shape_count": len(slide.shapes),
                "has_text": False,
                "has_images": False,
                "has_charts": False
            }
            
            for shape in slide.shapes:
                total_shapes += 1
                try:
                    if hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                        total_text_boxes += 1
                        slide_info["has_text"] = True
                    elif hasattr(shape, 'chart'):
                        total_charts += 1
                        slide_info["has_charts"] = True
                    elif hasattr(shape, 'image'):
                        total_images += 1
                        slide_info["has_images"] = True
                except:
                    pass
            
            slide_details.append(slide_info)
        
        # Get available slide layouts
        available_layouts = []
        for i, layout in enumerate(prs.slide_layouts):
            try:
                layout_name = layout.name if hasattr(layout, 'name') else f"Layout {i}"
                available_layouts.append({
                    "index": i,
                    "name": layout_name
                })
            except:
                available_layouts.append({
                    "index": i,
                    "name": f"Layout {i}"
                })
        
        return {
            "presentation_id": prs_id,
            "slide_count": len(prs.slides),
            "total_shapes": total_shapes,
            "content_summary": {
                "text_boxes": total_text_boxes,
                "images": total_images, 
                "charts": total_charts,
                "other_shapes": total_shapes - total_text_boxes - total_images - total_charts
            },
            "available_layouts": available_layouts,
            "slide_details": slide_details,
            "status": "ready"
        }
    
    def _post_process_slide(self, prs_id: str, slide_index: int):
        """Basic post-processing to fix common issues"""
        try:
            prs = self.presentations[prs_id]
            slide = prs.slides[slide_index]
            
            # Fix green rectangle fills on placeholder shapes
            for shape in slide.shapes:
                if hasattr(shape, 'fill'):
                    try:
                        if hasattr(shape.fill, 'solid'):
                            shape.fill.solid()
                            shape.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
                    except:
                        pass  # Skip if can't modify
                        
        except Exception as e:
            logger.warning(f"Post-processing failed for slide {slide_index}: {e}")

# =============================================================================
# MCP SERVER SETUP
# =============================================================================

# Initialize server and manager
server = Server("powerpoint-mcp-stable")
ppt_manager = StablePowerPointManager()

@server.list_tools()
async def handle_list_tools() -> List[Tool]:
    """List the core essential tools including deletion and file management capabilities"""
    return [
        Tool(
            name="create_presentation",
            description="Create a new PowerPoint presentation",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
                "additionalProperties": False
            }
        ),
        Tool(
            name="load_presentation",
            description="Load an existing PowerPoint presentation from file",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to the PowerPoint file to load"}
                },
                "required": ["file_path"],
                "additionalProperties": False,
                "examples": [
                    {
                        "file_path": "existing_presentation.pptx"
                    },
                    {
                        "file_path": "C:\\Documents\\quarterly_report.pptx"
                    }
                ]
            }
        ),
        Tool(
            name="add_slide",
            description="Add a new slide to a presentation with specified layout",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "layout_index": {"type": "integer", "description": "Slide layout index (0=title, 1=title+content, 6=blank, etc.)", "default": 6, "minimum": 0}
                },
                "required": ["presentation_id"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "layout_index": 1
                    },
                    {
                        "presentation_id": "ppt_0",
                        "layout_index": 6
                    }
                ]
            }
        ),
        Tool(
            name="add_text_box",
            description="Add a text box to a slide with comprehensive formatting options",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "text": {"type": "string", "description": "Text content"},
                    "left": {"type": "number", "default": 1, "description": "Left position in inches"},
                    "top": {"type": "number", "default": 1, "description": "Top position in inches"},
                    "width": {"type": "number", "default": 8, "description": "Width in inches"},
                    "height": {"type": "number", "default": 1, "description": "Height in inches"},
                    "font_size": {"type": "integer", "default": 18, "minimum": 8, "maximum": 72, "description": "Font size in points"},
                    "font_name": {"type": "string", "default": "Calibri", "description": "Font family name (e.g., Calibri, Arial, Times New Roman)"},
                    "font_color": {"type": "string", "description": "Font color - hex (#FF0000), RGB (255,0,0), or name (red, blue, etc.)"},
                    "bold": {"type": "boolean", "default": False, "description": "Bold text"},
                    "italic": {"type": "boolean", "default": False, "description": "Italic text"},
                    "underline": {"type": "boolean", "default": False, "description": "Underline text"},
                    "text_alignment": {"type": "string", "enum": ["left", "center", "right", "justify"], "default": "left", "description": "Text alignment"},
                    "fill_color": {"type": "string", "description": "Background fill color - hex (#FF0000), RGB (255,0,0), or name (red, blue, etc.)"},
                    "border_color": {"type": "string", "description": "Border color - hex (#FF0000), RGB (255,0,0), or name (red, blue, etc.)"},
                    "border_width": {"type": "number", "default": 0, "minimum": 0, "description": "Border width in points (0 = no border)"}
                },
                "required": ["presentation_id", "slide_index", "text"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 0,
                        "text": "Welcome to Our Presentation",
                        "font_size": 32,
                        "font_name": "Arial",
                        "font_color": "#0066CC",
                        "bold": True,
                        "text_alignment": "center"
                    },
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 1,
                        "text": "Key Points",
                        "font_color": "white",
                        "fill_color": "#4472C4",
                        "border_color": "black",
                        "border_width": 2
                    }
                ]
            }
        ),
        Tool(
            name="add_image",
            description="Add an image to a slide from URL or local file",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "image_source": {"type": "string", "description": "Image URL or local file path"},
                    "left": {"type": "number", "default": 1, "description": "Left position in inches"},
                    "top": {"type": "number", "default": 1, "description": "Top position in inches"},
                    "width": {"type": "number", "description": "Width in inches (optional)"},
                    "height": {"type": "number", "description": "Height in inches (optional)"}
                },
                "required": ["presentation_id", "slide_index", "image_source"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 1,
                        "image_source": "https://example.com/chart.png",
                        "width": 6,
                        "height": 4
                    }
                ]
            }
        ),
        Tool(
            name="add_chart",
            description="Add a chart to a slide with data series",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "chart_type": {"type": "string", "enum": ["column", "bar", "line", "pie", "area"]},
                    "categories": {"type": "array", "items": {"type": "string"}, "description": "Chart categories"},
                    "series_data": {
                        "type": "object",
                        "additionalProperties": {
                            "type": "array",
                            "items": {"type": "number"}
                        },
                        "description": "Series data as {series_name: [values]}"
                    },
                    "left": {"type": "number", "default": 2},
                    "top": {"type": "number", "default": 2},
                    "width": {"type": "number", "default": 6},
                    "height": {"type": "number", "default": 4.5}
                },
                "required": ["presentation_id", "slide_index", "chart_type", "categories", "series_data"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 2,
                        "chart_type": "column",
                        "categories": ["Q1", "Q2", "Q3", "Q4"],
                        "series_data": {
                            "Sales": [10, 15, 12, 18],
                            "Profit": [3, 5, 4, 7]
                        }
                    }
                ]
            }
        ),
        Tool(
            name="save_presentation",
            description="Save presentation to a file (handles both relative and absolute Windows paths)",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "file_path": {"type": "string", "description": "Output file path - can be relative (e.g., 'output/file.pptx') or absolute (e.g., 'C:\\Users\\name\\Documents\\file.pptx')"}
                },
                "required": ["presentation_id", "file_path"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "file_path": "my_presentation.pptx"
                    },
                    {
                        "presentation_id": "ppt_0", 
                        "file_path": "output/my_presentation.pptx"
                    },
                    {
                        "presentation_id": "ppt_0",
                        "file_path": "C:\\Users\\username\\Documents\\my_presentation.pptx"
                    }
                                 ]
             }
        ),
        Tool(
            name="extract_text",
            description="Extract all text content from a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"}
                },
                "required": ["presentation_id"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0"
                    }
                ]
            }
        ),
        Tool(
            name="get_presentation_info",
            description="Get comprehensive information about a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"}
                },
                "required": ["presentation_id"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0"
                    }
                ]
            }
        ),
        Tool(
            name="delete_shape",
            description="Delete a specific shape from a slide by index",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "shape_index": {"type": "integer", "description": "Shape index (0-based)", "minimum": 0}
                },
                "required": ["presentation_id", "slide_index", "shape_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 0,
                        "shape_index": 1
                    }
                ]
            }
        ),
        Tool(
            name="delete_slide",
            description="Delete an entire slide from the presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0}
                },
                "required": ["presentation_id", "slide_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 2
                    }
                ]
            }
        ),
        Tool(
            name="clear_slide",
            description="Clear all content from a slide but keep the slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0}
                },
                "required": ["presentation_id", "slide_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 1
                    }
                ]
            }
        ),
        Tool(
            name="list_slide_content",
            description="List all shapes on a slide to help with targeted deletion",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0}
                },
                "required": ["presentation_id", "slide_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 0
                    }
                ]
            }
        ),
        Tool(
            name="format_existing_text",
            description="Modify formatting of existing text shapes on slides",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "shape_index": {"type": "integer", "description": "Shape index (0-based)", "minimum": 0},
                    "font_size": {"type": "integer", "minimum": 8, "maximum": 72, "description": "Font size in points"},
                    "font_name": {"type": "string", "description": "Font family name (e.g., Arial, Calibri)"},
                    "font_color": {"type": "string", "description": "Font color - hex (#FF0000), RGB (255,0,0), or name (red, blue, etc.)"},
                    "bold": {"type": "boolean", "description": "Bold text"},
                    "italic": {"type": "boolean", "description": "Italic text"},
                    "underline": {"type": "boolean", "description": "Underline text"},
                    "text_alignment": {"type": "string", "enum": ["left", "center", "right", "justify"], "description": "Text alignment"}
                },
                "required": ["presentation_id", "slide_index", "shape_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 0,
                        "shape_index": 1,
                        "font_size": 24,
                        "font_color": "#FF0000",
                        "bold": True
                    }
                ]
            }
        ),
        Tool(
            name="set_slide_background",
            description="Set slide background color or image",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_index": {"type": "integer", "description": "Slide index (0-based)", "minimum": 0},
                    "background_color": {"type": "string", "description": "Background color - hex (#FF0000), RGB (255,0,0), or name (red, blue, etc.)"},
                    "background_image": {"type": "string", "description": "Background image URL or local file path"}
                },
                "required": ["presentation_id", "slide_index"],
                "additionalProperties": False,
                "examples": [
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 0,
                        "background_color": "#E6F3FF"
                    },
                    {
                        "presentation_id": "ppt_0",
                        "slide_index": 1,
                        "background_image": "https://example.com/background.jpg"
                    }
                ]
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
    """Handle tool calls with validation and enhanced feedback"""
    
    try:
        # Validate arguments
        validated_args = validate_basic_args(name, arguments)
        
        if name == "create_presentation":
            prs_id = ppt_manager.create_presentation()
            message = format_success_message(name, presentation_id=prs_id)
            return [TextContent(type="text", text=message)]
        
        elif name == "load_presentation":
            file_path = validated_args["file_path"]
            
            prs_id = ppt_manager.load_presentation(file_path)
            
            # Get info for success message
            prs = ppt_manager.presentations[prs_id]
            slide_count = len(prs.slides)
            
            message = format_success_message(
                name, presentation_id=prs_id, file_path=file_path, slide_count=slide_count
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "add_slide":
            prs_id = validated_args["presentation_id"]
            layout_index = validated_args.get("layout_index", 6)
            
            slide_index = ppt_manager.add_slide(prs_id, layout_index)
            
            # Get info for success message
            prs = ppt_manager.presentations[prs_id]
            total_slides = len(prs.slides)
            
            # Try to get layout name
            try:
                layout_name = prs.slide_layouts[layout_index].name
            except:
                layout_name = f"Layout {layout_index}"
            
            message = format_success_message(
                name, slide_index=slide_index, layout_index=layout_index, 
                layout_name=layout_name, total_slides=total_slides
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "add_text_box":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            text = validated_args["text"]
            left = validated_args.get("left", 1)
            top = validated_args.get("top", 1)
            width = validated_args.get("width", 8)
            height = validated_args.get("height", 1)
            font_size = validated_args.get("font_size", 18)
            font_name = validated_args.get("font_name", "Calibri")
            font_color = validated_args.get("font_color")
            bold = validated_args.get("bold", False)
            italic = validated_args.get("italic", False)
            underline = validated_args.get("underline", False)
            text_alignment = validated_args.get("text_alignment", "left")
            fill_color = validated_args.get("fill_color")
            border_color = validated_args.get("border_color")
            border_width = validated_args.get("border_width", 0)
            
            success = ppt_manager.add_text_box(
                prs_id, slide_index, text, left, top, width, height, font_size, font_name, font_color,
                bold, italic, underline, text_alignment, fill_color, border_color, border_width
            )
            
            message = format_success_message(
                name, slide_index=slide_index, font_size=font_size, font_name=font_name, 
                text_alignment=text_alignment, font_color=font_color, fill_color=fill_color, text=text
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "add_image":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            image_source = validated_args["image_source"]
            left = validated_args.get("left", 1)
            top = validated_args.get("top", 1)
            width = validated_args.get("width")
            height = validated_args.get("height")
            
            success = ppt_manager.add_image(
                prs_id, slide_index, image_source, left, top, width, height
            )
            
            message = format_success_message(
                name, slide_index=slide_index, image_source=image_source
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "add_chart":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            chart_type = validated_args["chart_type"]
            categories = validated_args["categories"]
            series_data = validated_args["series_data"]
            left = validated_args.get("left", 2)
            top = validated_args.get("top", 2)
            width = validated_args.get("width", 6)
            height = validated_args.get("height", 4.5)
            
            success = ppt_manager.add_chart(
                prs_id, slide_index, chart_type, categories, series_data, left, top, width, height
            )
            
            message = format_success_message(
                name, slide_index=slide_index, chart_type=chart_type, 
                categories=categories, series_data=series_data
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "save_presentation":
            prs_id = validated_args["presentation_id"]
            file_path = validated_args["file_path"]
            
            saved_path = ppt_manager.save_presentation(prs_id, file_path)
            
            # Return with embedded resource for immediate access
            try:
                with open(saved_path, 'rb') as f:
                    file_data = f.read()
                
                message = format_success_message(name, file_path=saved_path)
                
                return [
                    TextContent(type="text", text=message),
                    EmbeddedResource(
                        type="resource",
                        resource=EmbeddedResource(
                            uri=f"file://{saved_path}",
                            mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            data=file_data
                        )
                    )
                ]
            except Exception as e:
                # Fallback to text message only
                message = format_success_message(name, file_path=saved_path)
                return [TextContent(type="text", text=message)]
        
        elif name == "extract_text":
            prs_id = validated_args["presentation_id"]
            
            extracted_text = ppt_manager.extract_text(prs_id)
            
            # Count total text items
            text_items = sum(len(slide["text_content"]) for slide in extracted_text)
            slide_count = len(extracted_text)
            
            message = format_success_message(
                name, slide_count=slide_count, text_items=text_items
            )
            
            # Format the extracted text for display
            if extracted_text:
                text_summary = []
                for slide in extracted_text:
                    if slide["text_content"]:
                        slide_text = f"Slide {slide['slide_number']}:"
                        for item in slide["text_content"]:
                            preview = item["text"][:80] + "..." if len(item["text"]) > 80 else item["text"]
                            slide_text += f"\n  - {preview}"
                        text_summary.append(slide_text)
                
                if text_summary:
                    full_message = f"{message}\n\n" + "\n\n".join(text_summary)
                else:
                    full_message = f"{message}\n(No text content found)"
            else:
                full_message = f"{message}\n(No slides found)"
            
            return [TextContent(type="text", text=full_message)]
        
        elif name == "get_presentation_info":
            prs_id = validated_args["presentation_id"]
            
            info = ppt_manager.get_presentation_info(prs_id)
            
            message = format_success_message(
                name, slide_count=info["slide_count"], total_shapes=info["total_shapes"]
            )
            
            # Format comprehensive info
            content_summary = info["content_summary"]
            available_layouts = info["available_layouts"]
            
            info_details = f"""📊 Content Summary:
  • Text boxes: {content_summary['text_boxes']}
  • Images: {content_summary['images']}
  • Charts: {content_summary['charts']}
  • Other shapes: {content_summary['other_shapes']}

🎨 Available Layouts:"""
            
            for layout in available_layouts[:5]:  # Show first 5 layouts
                info_details += f"\n  [{layout['index']}] {layout['name']}"
            
            if len(available_layouts) > 5:
                info_details += f"\n  ... and {len(available_layouts) - 5} more layouts"
            
            full_message = f"{message}\n\n{info_details}"
            return [TextContent(type="text", text=full_message)]
        
        elif name == "delete_shape":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            shape_index = validated_args["shape_index"]
            
            # Get shape type before deletion for better message
            prs = ppt_manager.presentations[prs_id]
            slide = prs.slides[slide_index]
            shape = slide.shapes[shape_index]
            shape_type = "shape"
            try:
                if hasattr(shape, 'text_frame'):
                    shape_type = "text box"
                elif hasattr(shape, 'chart'):
                    shape_type = "chart"
                elif hasattr(shape, 'image'):
                    shape_type = "image"
            except:
                pass
            
            success = ppt_manager.delete_shape(prs_id, slide_index, shape_index)
            message = format_success_message(
                name, slide_index=slide_index, shape_index=shape_index, shape_type=shape_type
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "delete_slide":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            
            # Get slide count before deletion
            prs = ppt_manager.presentations[prs_id]
            original_count = len(prs.slides)
            
            success = ppt_manager.delete_slide(prs_id, slide_index)
            remaining_slides = original_count - 1
            message = format_success_message(
                name, slide_index=slide_index, remaining_slides=remaining_slides
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "clear_slide":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            
            # Get shape count before clearing
            prs = ppt_manager.presentations[prs_id]
            slide = prs.slides[slide_index]
            shapes_cleared = len(slide.shapes)
            
            success = ppt_manager.clear_slide(prs_id, slide_index)
            message = format_success_message(
                name, slide_index=slide_index, shapes_cleared=shapes_cleared
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "list_slide_content":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            
            content = ppt_manager.list_slide_content(prs_id, slide_index)
            message = format_success_message(
                name, slide_index=slide_index, shape_count=content["shape_count"]
            )
            
            # Format the detailed content list
            if content["shapes"]:
                content_list = "\n".join([
                    f"  [{shape['index']}] {shape['type']}: {shape['description']}"
                    for shape in content["shapes"]
                ])
                full_message = f"{message}\n{content_list}"
            else:
                full_message = f"{message}\n  (No shapes on this slide)"
            
            return [TextContent(type="text", text=full_message)]
        
        elif name == "format_existing_text":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            shape_index = validated_args["shape_index"]
            font_size = validated_args.get("font_size")
            font_name = validated_args.get("font_name")
            font_color = validated_args.get("font_color")
            bold = validated_args.get("bold")
            italic = validated_args.get("italic")
            underline = validated_args.get("underline")
            text_alignment = validated_args.get("text_alignment")
            
            success = ppt_manager.format_existing_text(
                prs_id, slide_index, shape_index, font_size, font_name, font_color,
                bold, italic, underline, text_alignment
            )
            
            message = format_success_message(
                name, slide_index=slide_index, shape_index=shape_index,
                font_size=font_size, font_name=font_name, font_color=font_color, text_alignment=text_alignment
            )
            return [TextContent(type="text", text=message)]
        
        elif name == "set_slide_background":
            prs_id = validated_args["presentation_id"]
            slide_index = validated_args["slide_index"]
            background_color = validated_args.get("background_color")
            background_image = validated_args.get("background_image")
            
            success = ppt_manager.set_slide_background(
                prs_id, slide_index, background_color, background_image
            )
            
            message = format_success_message(
                name, slide_index=slide_index, background_color=background_color, background_image=background_image
            )
            return [TextContent(type="text", text=message)]
        
        else:
            raise ValueError(f"Unknown tool: {name}")
    
    except Exception as e:
        logger.error(f"Error in tool {name}: {e}")
        return [TextContent(
            type="text",
            text=f"Error: {str(e)}"
        )]

async def main():
    """Main entry point for the stable PowerPoint MCP server"""
    options = InitializationOptions(
        server_name="powerpoint-mcp-stable",
        server_version="1.0.0",
        capabilities={
            "resources": {},
            "tools": {}
        }
    )
    
    try:
        async with stdio_server() as (read_stream, write_stream):
            await server.run(read_stream, write_stream, options)
    except KeyboardInterrupt:
        logger.info("Server interrupted by user")
    except Exception as e:
        logger.error(f"Server error: {e}")

if __name__ == "__main__":
    asyncio.run(main()) 