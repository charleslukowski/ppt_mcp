#!/usr/bin/env python3
"""
Basic Usage Examples for PowerPoint MCP Server

This script demonstrates how to interact with the PowerPoint MCP server
to create and manipulate PowerPoint presentations programmatically.
"""

import json
import asyncio
from typing import Dict, Any

# Note: In a real implementation, you would use the MCP client library
# This is a simplified example showing the structure of MCP calls

class MockMCPClient:
    """Mock MCP client for demonstration purposes"""
    
    def __init__(self):
        # In real usage, this would connect to the actual MCP server
        pass
    
    async def call_tool(self, name: str, arguments: Dict[str, Any]) -> str:
        """Simulate calling an MCP tool"""
        print(f"Calling tool: {name}")
        print(f"Arguments: {json.dumps(arguments, indent=2)}")
        
        # Mock responses for demonstration
        if name == "create_presentation":
            return "Created presentation with ID: prs_0"
        elif name == "add_slide":
            return "Added slide 0 to presentation prs_0"
        elif name == "add_text_box":
            return f"Added text box to slide {arguments['slide_index']}"
        elif name == "save_presentation":
            return f"Saved presentation {arguments['presentation_id']} to {arguments['file_path']}"
        else:
            return f"Tool {name} executed successfully"

async def example_1_simple_presentation():
    """Example 1: Create a simple presentation with title and content"""
    print("\n=== Example 1: Simple Presentation ===")
    
    client = MockMCPClient()
    
    # 1. Create a new presentation
    result = await client.call_tool("create_presentation", {})
    print(f"Result: {result}")
    
    # 2. Add a title slide
    result = await client.call_tool("add_slide", {
        "presentation_id": "prs_0",
        "layout_index": 0  # Title slide layout
    })
    print(f"Result: {result}")
    
    # 3. Add title text
    result = await client.call_tool("add_text_box", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "text": "Welcome to Our Company",
        "left": 1,
        "top": 1,
        "width": 8,
        "height": 1.5,
        "font_size": 32,
        "bold": True
    })
    print(f"Result: {result}")
    
    # 4. Add subtitle
    result = await client.call_tool("add_text_box", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "text": "Annual Report 2024",
        "left": 1,
        "top": 3,
        "width": 8,
        "height": 1,
        "font_size": 24,
        "italic": True
    })
    print(f"Result: {result}")
    
    # 5. Save the presentation
    result = await client.call_tool("save_presentation", {
        "presentation_id": "prs_0",
        "file_path": "welcome_presentation.pptx"
    })
    print(f"Result: {result}")

async def example_2_data_driven_presentation():
    """Example 2: Create a presentation with charts and data"""
    print("\n=== Example 2: Data-Driven Presentation ===")
    
    client = MockMCPClient()
    
    # 1. Create presentation from template
    result = await client.call_tool("create_presentation", {
        "template_path": "company_template.pptx"
    })
    print(f"Result: {result}")
    
    # 2. Add a content slide
    result = await client.call_tool("add_slide", {
        "presentation_id": "prs_0",
        "layout_index": 6  # Blank layout for custom content
    })
    print(f"Result: {result}")
    
    # 3. Add title
    result = await client.call_tool("add_text_box", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "text": "Q4 Sales Performance",
        "left": 1,
        "top": 0.5,
        "width": 8,
        "height": 1,
        "font_size": 28,
        "bold": True
    })
    print(f"Result: {result}")
    
    # 4. Add a chart with sales data
    result = await client.call_tool("add_chart", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "chart_type": "column",
        "categories": ["October", "November", "December"],
        "series_data": {
            "Revenue": [150000, 180000, 220000],
            "Target": [160000, 170000, 200000]
        },
        "left": 1,
        "top": 2,
        "width": 8,
        "height": 4
    })
    print(f"Result: {result}")
    
    # 5. Save the presentation
    result = await client.call_tool("save_presentation", {
        "presentation_id": "prs_0",
        "file_path": "q4_sales_report.pptx"
    })
    print(f"Result: {result}")

async def example_3_image_integration():
    """Example 3: Create a presentation with images from URLs"""
    print("\n=== Example 3: Image Integration ===")
    
    client = MockMCPClient()
    
    # 1. Create new presentation
    result = await client.call_tool("create_presentation", {})
    print(f"Result: {result}")
    
    # 2. Add slide
    result = await client.call_tool("add_slide", {
        "presentation_id": "prs_0",
        "layout_index": 6
    })
    print(f"Result: {result}")
    
    # 3. Add title
    result = await client.call_tool("add_text_box", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "text": "Product Showcase",
        "left": 1,
        "top": 0.5,
        "width": 8,
        "height": 1,
        "font_size": 32,
        "bold": True
    })
    print(f"Result: {result}")
    
    # 4. Add image from URL
    result = await client.call_tool("add_image", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "image_source": "https://picsum.photos/800/600",
        "left": 1,
        "top": 2,
        "width": 6,
        "height": 4
    })
    print(f"Result: {result}")
    
    # 5. Add description text
    result = await client.call_tool("add_text_box", {
        "presentation_id": "prs_0",
        "slide_index": 0,
        "text": "Our latest product features innovative design and cutting-edge technology.",
        "left": 7.5,
        "top": 2,
        "width": 2,
        "height": 4,
        "font_size": 16
    })
    print(f"Result: {result}")
    
    # 6. Save presentation
    result = await client.call_tool("save_presentation", {
        "presentation_id": "prs_0",
        "file_path": "product_showcase.pptx"
    })
    print(f"Result: {result}")

async def example_4_json_driven_presentation():
    """Example 4: Create presentation from JSON schema"""
    print("\n=== Example 4: JSON-Driven Presentation ===")
    
    client = MockMCPClient()
    
    # Define presentation structure in JSON
    presentation_data = {
        "intro_slide": {
            "title": "Company Overview",
            "content": "Leading provider of innovative solutions since 2010"
        },
        "mission_slide": {
            "title": "Our Mission",
            "content": "To deliver exceptional value through technology and innovation"
        },
        "team_slide": {
            "title": "Our Team",
            "content": "50+ professionals across engineering, design, and business"
        },
        "contact_slide": {
            "title": "Get In Touch",
            "content": "Email: contact@company.com\nPhone: (555) 123-4567"
        }
    }
    
    # Create presentation from JSON
    result = await client.call_tool("create_from_json", {
        "json_data": presentation_data,
        "template_path": "corporate_template.pptx"
    })
    print(f"Result: {result}")
    
    # Save the generated presentation
    result = await client.call_tool("save_presentation", {
        "presentation_id": "prs_0",
        "file_path": "company_overview.pptx"
    })
    print(f"Result: {result}")

async def example_5_content_extraction():
    """Example 5: Extract content from existing presentation"""
    print("\n=== Example 5: Content Extraction ===")
    
    client = MockMCPClient()
    
    # 1. Load existing presentation
    result = await client.call_tool("load_presentation", {
        "file_path": "existing_presentation.pptx"
    })
    print(f"Result: {result}")
    
    # 2. Get presentation info
    result = await client.call_tool("get_presentation_info", {
        "presentation_id": "prs_0"
    })
    print(f"Result: {result}")
    
    # 3. Extract all text content
    result = await client.call_tool("extract_text", {
        "presentation_id": "prs_0"
    })
    print(f"Result: {result}")

async def main():
    """Run all examples"""
    print("PowerPoint MCP Server - Usage Examples")
    print("=" * 50)
    
    await example_1_simple_presentation()
    await example_2_data_driven_presentation()
    await example_3_image_integration()
    await example_4_json_driven_presentation()
    await example_5_content_extraction()
    
    print("\n" + "=" * 50)
    print("All examples completed!")
    print("\nNote: These are mock examples showing the structure of MCP calls.")
    print("In real usage, you would connect to the actual PowerPoint MCP server.")

if __name__ == "__main__":
    asyncio.run(main()) 