#!/usr/bin/env python3
"""
PowerPoint MCP Server

A Model Context Protocol server for PowerPoint presentation manipulation using python-pptx.
Provides tools for creating, editing, and managing PowerPoint presentations programmatically.

Based on research of existing MCP servers and python-pptx best practices.
Optimized for use with Cursor IDE and AI-assisted development workflows.
"""

import asyncio
import json
import logging
import os
import sys
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from urllib.request import urlopen
from urllib.parse import urlparse

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError as e:
    print(f"python-pptx library not found: {e}")
    print("Please install with: pip install python-pptx")
    sys.exit(1)

try:
    from mcp.server import Server
    from mcp.server.models import InitializationOptions
    from mcp.server.stdio import stdio_server
    from mcp.types import (
        Resource,
        Tool,
        TextContent,
        ImageContent,
        EmbeddedResource,
        LoggingLevel
    )
except ImportError as e:
    print(f"MCP library not found: {e}")
    print("Please install with: pip install mcp")
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("powerpoint-mcp-server")

# Server instance
server = Server("powerpoint-mcp-server")

class PowerPointManager:
    """Manages PowerPoint presentation operations"""
    
    def __init__(self):
        self.presentations: Dict[str, Presentation] = {}
        self.temp_files: List[str] = []
    
    def create_presentation(self, template_path: Optional[str] = None) -> str:
        """Create a new presentation, optionally from a template"""
        try:
            if template_path and os.path.exists(template_path):
                prs = Presentation(template_path)
                logger.info(f"Created presentation from template: {template_path}")
            else:
                prs = Presentation()
                logger.info("Created new blank presentation")
            
            # Generate unique ID for this presentation
            prs_id = f"prs_{len(self.presentations)}"
            self.presentations[prs_id] = prs
            return prs_id
        except Exception as e:
            logger.error(f"Error creating presentation: {e}")
            raise
    
    def load_presentation(self, file_path: str) -> str:
        """Load an existing presentation"""
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Presentation file not found: {file_path}")
            
            prs = Presentation(file_path)
            prs_id = f"prs_{len(self.presentations)}"
            self.presentations[prs_id] = prs
            logger.info(f"Loaded presentation: {file_path}")
            return prs_id
        except Exception as e:
            logger.error(f"Error loading presentation: {e}")
            raise
    
    def save_presentation(self, prs_id: str, file_path: str) -> bool:
        """Save a presentation to file"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            self.presentations[prs_id].save(file_path)
            logger.info(f"Saved presentation {prs_id} to {file_path}")
            return True
        except Exception as e:
            logger.error(f"Error saving presentation: {e}")
            raise
    
    def add_slide(self, prs_id: str, layout_index: int = 6) -> int:
        """Add a new slide to the presentation"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            slide_layout = prs.slide_layouts[layout_index]
            slide = prs.slides.add_slide(slide_layout)
            slide_index = len(prs.slides) - 1
            
            logger.info(f"Added slide {slide_index} to presentation {prs_id}")
            return slide_index
        except Exception as e:
            logger.error(f"Error adding slide: {e}")
            raise
    
    def add_text_box(self, prs_id: str, slide_index: int, text: str, 
                     left: float = 1, top: float = 1, width: float = 8, height: float = 1,
                     font_size: int = 18, bold: bool = False, italic: bool = False) -> bool:
        """Add a text box to a slide"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            if slide_index >= len(prs.slides):
                raise ValueError(f"Slide index {slide_index} out of range")
            
            slide = prs.slides[slide_index]
            txt_box = slide.shapes.add_textbox(
                Inches(left), Inches(top), Inches(width), Inches(height)
            )
            txt_frame = txt_box.text_frame
            txt_frame.text = text
            
            # Apply formatting
            for paragraph in txt_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
                    run.font.bold = bold
                    run.font.italic = italic
            
            logger.info(f"Added text box to slide {slide_index} in presentation {prs_id}")
            return True
        except Exception as e:
            logger.error(f"Error adding text box: {e}")
            raise
    
    def add_image(self, prs_id: str, slide_index: int, image_source: str,
                  left: float = 1, top: float = 1, width: Optional[float] = None, height: Optional[float] = None) -> bool:
        """Add an image to a slide from file path or URL"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            if slide_index >= len(prs.slides):
                raise ValueError(f"Slide index {slide_index} out of range")
            
            slide = prs.slides[slide_index]
            
            # Handle URL or file path
            if image_source.startswith(('http://', 'https://')):
                # Download image from URL
                image_data = BytesIO(urlopen(image_source).read())
                if width and height:
                    slide.shapes.add_picture(image_data, Inches(left), Inches(top), Inches(width), Inches(height))
                else:
                    slide.shapes.add_picture(image_data, Inches(left), Inches(top))
            else:
                # Load from file path
                if not os.path.exists(image_source):
                    raise FileNotFoundError(f"Image file not found: {image_source}")
                
                if width and height:
                    slide.shapes.add_picture(image_source, Inches(left), Inches(top), Inches(width), Inches(height))
                else:
                    slide.shapes.add_picture(image_source, Inches(left), Inches(top))
            
            logger.info(f"Added image to slide {slide_index} in presentation {prs_id}")
            return True
        except Exception as e:
            logger.error(f"Error adding image: {e}")
            raise
    
    def add_chart(self, prs_id: str, slide_index: int, chart_type: str, 
                  categories: List[str], series_data: Dict[str, List[float]],
                  left: float = 2, top: float = 2, width: float = 6, height: float = 4.5) -> bool:
        """Add a chart to a slide"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            if slide_index >= len(prs.slides):
                raise ValueError(f"Slide index {slide_index} out of range")
            
            slide = prs.slides[slide_index]
            
            # Create chart data
            chart_data = CategoryChartData()
            chart_data.categories = categories
            
            for series_name, data in series_data.items():
                chart_data.add_series(series_name, data)
            
            # Map chart type string to enum
            chart_type_map = {
                'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
                'bar': XL_CHART_TYPE.BAR_CLUSTERED,
                'line': XL_CHART_TYPE.LINE,
                'pie': XL_CHART_TYPE.PIE
            }
            
            chart_type_enum = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
            
            slide.shapes.add_chart(
                chart_type_enum,
                Inches(left), Inches(top), Inches(width), Inches(height),
                chart_data
            )
            
            logger.info(f"Added {chart_type} chart to slide {slide_index} in presentation {prs_id}")
            return True
        except Exception as e:
            logger.error(f"Error adding chart: {e}")
            raise
    
    def extract_text(self, prs_id: str) -> List[Dict[str, Any]]:
        """Extract all text content from a presentation"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            extracted_content = []
            
            for slide_idx, slide in enumerate(prs.slides):
                slide_content = {
                    'slide_index': slide_idx,
                    'text_content': []
                }
                
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            text = ''.join(run.text for run in paragraph.runs)
                            if text.strip():
                                slide_content['text_content'].append(text.strip())
                
                extracted_content.append(slide_content)
            
            logger.info(f"Extracted text from presentation {prs_id}")
            return extracted_content
        except Exception as e:
            logger.error(f"Error extracting text: {e}")
            raise
    
    def get_presentation_info(self, prs_id: str) -> Dict[str, Any]:
        """Get information about a presentation"""
        try:
            if prs_id not in self.presentations:
                raise ValueError(f"Presentation {prs_id} not found")
            
            prs = self.presentations[prs_id]
            
            info = {
                'presentation_id': prs_id,
                'slide_count': len(prs.slides),
                'slide_layouts_count': len(prs.slide_layouts),
                'slide_master_count': len(prs.slide_masters)
            }
            
            return info
        except Exception as e:
            logger.error(f"Error getting presentation info: {e}")
            raise
    
    def cleanup(self):
        """Clean up temporary files"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except Exception as e:
                logger.warning(f"Could not remove temp file {temp_file}: {e}")
        self.temp_files.clear()

# Global PowerPoint manager instance
ppt_manager = PowerPointManager()

@server.list_resources()
async def handle_list_resources() -> List[Resource]:
    """List available PowerPoint presentations as resources"""
    resources = []
    
    for prs_id, prs in ppt_manager.presentations.items():
        resources.append(
            Resource(
                uri=f"powerpoint://{prs_id}",
                name=f"PowerPoint Presentation {prs_id}",
                description=f"PowerPoint presentation with {len(prs.slides)} slides",
                mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        )
    
    return resources

@server.read_resource()
async def handle_read_resource(uri: str) -> str:
    """Read a PowerPoint presentation resource"""
    if not uri.startswith("powerpoint://"):
        raise ValueError(f"Unsupported URI scheme: {uri}")
    
    prs_id = uri.replace("powerpoint://", "")
    
    if prs_id not in ppt_manager.presentations:
        raise ValueError(f"Presentation {prs_id} not found")
    
    # Extract and return presentation information
    info = ppt_manager.get_presentation_info(prs_id)
    text_content = ppt_manager.extract_text(prs_id)
    
    result = {
        'presentation_info': info,
        'content': text_content
    }
    
    return json.dumps(result, indent=2)

@server.list_tools()
async def handle_list_tools() -> List[Tool]:
    """List available PowerPoint manipulation tools"""
    return [
        Tool(
            name="create_presentation",
            description="Create a new PowerPoint presentation, optionally from a template",
            inputSchema={
                "type": "object",
                "properties": {
                    "template_path": {
                        "type": "string",
                        "description": "Optional path to template file"
                    }
                }
            }
        ),
        Tool(
            name="load_presentation",
            description="Load an existing PowerPoint presentation from file",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the PowerPoint file to load"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="save_presentation",
            description="Save a PowerPoint presentation to file",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {
                        "type": "string",
                        "description": "ID of the presentation to save"
                    },
                    "file_path": {
                        "type": "string",
                        "description": "Path where to save the presentation"
                    }
                },
                "required": ["presentation_id", "file_path"]
            }
        ),
        Tool(
            name="add_slide",
            description="Add a new slide to a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {
                        "type": "string",
                        "description": "ID of the presentation"
                    },
                    "layout_index": {
                        "type": "integer",
                        "description": "Slide layout index (0=title, 1=title+content, 6=blank, etc.)",
                        "default": 6
                    }
                },
                "required": ["presentation_id"]
            }
        ),
        Tool(
            name="add_text_box",
            description="Add a text box to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string"},
                    "slide_index": {"type": "integer"},
                    "text": {"type": "string"},
                    "left": {"type": "number", "default": 1},
                    "top": {"type": "number", "default": 1},
                    "width": {"type": "number", "default": 8},
                    "height": {"type": "number", "default": 1},
                    "font_size": {"type": "integer", "default": 18},
                    "bold": {"type": "boolean", "default": False},
                    "italic": {"type": "boolean", "default": False}
                },
                "required": ["presentation_id", "slide_index", "text"]
            }
        ),
        Tool(
            name="add_image",
            description="Add an image to a slide from file path or URL",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string"},
                    "slide_index": {"type": "integer"},
                    "image_source": {"type": "string", "description": "File path or URL to image"},
                    "left": {"type": "number", "default": 1},
                    "top": {"type": "number", "default": 1},
                    "width": {"type": "number"},
                    "height": {"type": "number"}
                },
                "required": ["presentation_id", "slide_index", "image_source"]
            }
        ),
        Tool(
            name="add_chart",
            description="Add a chart to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string"},
                    "slide_index": {"type": "integer"},
                    "chart_type": {"type": "string", "enum": ["column", "bar", "line", "pie"]},
                    "categories": {"type": "array", "items": {"type": "string"}},
                    "series_data": {
                        "type": "object",
                        "description": "Dictionary with series names as keys and data arrays as values"
                    },
                    "left": {"type": "number", "default": 2},
                    "top": {"type": "number", "default": 2},
                    "width": {"type": "number", "default": 6},
                    "height": {"type": "number", "default": 4.5}
                },
                "required": ["presentation_id", "slide_index", "chart_type", "categories", "series_data"]
            }
        ),
        Tool(
            name="extract_text",
            description="Extract all text content from a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string"}
                },
                "required": ["presentation_id"]
            }
        ),
        Tool(
            name="get_presentation_info",
            description="Get information about a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string"}
                },
                "required": ["presentation_id"]
            }
        ),
        Tool(
            name="create_from_json",
            description="Create a presentation from structured JSON data (schema-driven approach)",
            inputSchema={
                "type": "object",
                "properties": {
                    "json_data": {
                        "type": "object",
                        "description": "Structured presentation data"
                    },
                    "template_path": {
                        "type": "string",
                        "description": "Optional template file path"
                    }
                },
                "required": ["json_data"]
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
    """Handle tool calls for PowerPoint operations"""
    try:
        if name == "create_presentation":
            template_path = arguments.get("template_path")
            prs_id = ppt_manager.create_presentation(template_path)
            return [TextContent(
                type="text",
                text=f"Created presentation with ID: {prs_id}"
            )]
        
        elif name == "load_presentation":
            file_path = arguments["file_path"]
            prs_id = ppt_manager.load_presentation(file_path)
            return [TextContent(
                type="text",
                text=f"Loaded presentation with ID: {prs_id}"
            )]
        
        elif name == "save_presentation":
            prs_id = arguments["presentation_id"]
            file_path = arguments["file_path"]
            success = ppt_manager.save_presentation(prs_id, file_path)
            return [TextContent(
                type="text",
                text=f"Saved presentation {prs_id} to {file_path}"
            )]
        
        elif name == "add_slide":
            prs_id = arguments["presentation_id"]
            layout_index = arguments.get("layout_index", 6)
            slide_index = ppt_manager.add_slide(prs_id, layout_index)
            return [TextContent(
                type="text",
                text=f"Added slide {slide_index} to presentation {prs_id}"
            )]
        
        elif name == "add_text_box":
            result = ppt_manager.add_text_box(
                arguments["presentation_id"],
                arguments["slide_index"],
                arguments["text"],
                arguments.get("left", 1),
                arguments.get("top", 1),
                arguments.get("width", 8),
                arguments.get("height", 1),
                arguments.get("font_size", 18),
                arguments.get("bold", False),
                arguments.get("italic", False)
            )
            return [TextContent(
                type="text",
                text=f"Added text box to slide {arguments['slide_index']}"
            )]
        
        elif name == "add_image":
            result = ppt_manager.add_image(
                arguments["presentation_id"],
                arguments["slide_index"],
                arguments["image_source"],
                arguments.get("left", 1),
                arguments.get("top", 1),
                arguments.get("width"),
                arguments.get("height")
            )
            return [TextContent(
                type="text",
                text=f"Added image to slide {arguments['slide_index']}"
            )]
        
        elif name == "add_chart":
            result = ppt_manager.add_chart(
                arguments["presentation_id"],
                arguments["slide_index"],
                arguments["chart_type"],
                arguments["categories"],
                arguments["series_data"],
                arguments.get("left", 2),
                arguments.get("top", 2),
                arguments.get("width", 6),
                arguments.get("height", 4.5)
            )
            return [TextContent(
                type="text",
                text=f"Added {arguments['chart_type']} chart to slide {arguments['slide_index']}"
            )]
        
        elif name == "extract_text":
            prs_id = arguments["presentation_id"]
            content = ppt_manager.extract_text(prs_id)
            return [TextContent(
                type="text",
                text=json.dumps(content, indent=2)
            )]
        
        elif name == "get_presentation_info":
            prs_id = arguments["presentation_id"]
            info = ppt_manager.get_presentation_info(prs_id)
            return [TextContent(
                type="text",
                text=json.dumps(info, indent=2)
            )]
        
        elif name == "create_from_json":
            json_data = arguments["json_data"]
            template_path = arguments.get("template_path")
            
            # Create presentation from JSON schema (simplified implementation)
            prs_id = ppt_manager.create_presentation(template_path)
            
            # Process JSON data to create slides
            # This is a simplified version - could be expanded based on specific schema
            if isinstance(json_data, dict):
                for slide_name, slide_data in json_data.items():
                    slide_index = ppt_manager.add_slide(prs_id)
                    
                    if "title" in slide_data:
                        ppt_manager.add_text_box(prs_id, slide_index, slide_data["title"], 
                                               top=1, font_size=24, bold=True)
                    
                    if "content" in slide_data:
                        ppt_manager.add_text_box(prs_id, slide_index, slide_data["content"], 
                                               top=2.5, height=4)
            
            return [TextContent(
                type="text",
                text=f"Created presentation from JSON data with ID: {prs_id}"
            )]
        
        else:
            raise ValueError(f"Unknown tool: {name}")
    
    except Exception as e:
        logger.error(f"Error in tool {name}: {e}")
        return [TextContent(
            type="text",
            text=f"Error: {str(e)}"
        )]

async def main():
    """Main entry point for the PowerPoint MCP server"""
    # Initialize server options
    options = InitializationOptions(
        server_name="powerpoint-mcp-server",
        server_version="1.0.0",
        capabilities={
            "resources": {},
            "tools": {},
            "logging": {}
        }
    )
    
    try:
        async with stdio_server() as (read_stream, write_stream):
            await server.run(
                read_stream,
                write_stream,
                options
            )
    except KeyboardInterrupt:
        logger.info("Server interrupted by user")
    except Exception as e:
        logger.error(f"Server error: {e}")
    finally:
        # Cleanup
        ppt_manager.cleanup()
        logger.info("PowerPoint MCP server shutdown complete")

if __name__ == "__main__":
    asyncio.run(main()) 