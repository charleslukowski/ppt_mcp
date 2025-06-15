#!/usr/bin/env python3
"""
Test script for PowerPoint screenshot functionality

This script demonstrates the new screenshot feature that captures each slide
of a PowerPoint presentation as an image file for vision analysis.

Note: This feature is Windows-only and requires Microsoft PowerPoint to be installed.
"""

import asyncio
import json
import os
import tempfile
import platform
from pathlib import Path

# Check if we're on Windows
if platform.system() != "Windows":
    print("❌ Screenshot feature is only available on Windows")
    exit(1)

try:
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
except ImportError:
    print("❌ MCP client library not found. Please install with: pip install mcp")
    exit(1)


async def test_screenshot_functionality():
    """Test the screenshot functionality"""
    
    print("🖼️  Testing PowerPoint Screenshot Functionality")
    print("=" * 50)
    
    # Start the MCP server
    server_params = StdioServerParameters(
        command="python", 
        args=["powerpoint_mcp_server.py"]
    )
    
    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            # Initialize
            await session.initialize()
            
            # Test 1: Check if screenshot tool is available
            print("\n📋 Step 1: Checking available tools...")
            tools_result = await session.list_tools()
            screenshot_tool = None
            
            for tool in tools_result.tools:
                if tool.name == "screenshot_slides":
                    screenshot_tool = tool
                    break
            
            if screenshot_tool:
                print("✅ Screenshot tool found!")
                print(f"   Description: {screenshot_tool.description}")
            else:
                print("❌ Screenshot tool not found in available tools")
                return False
            
            # Test 2: Create a sample presentation for testing
            print("\n📋 Step 2: Creating sample presentation...")
            
            # Create presentation
            create_result = await session.call_tool("create_presentation", {})
            prs_id = create_result.content[0].text.split(": ")[1]
            print(f"✅ Created presentation: {prs_id}")
            
            # Add some sample slides
            for i in range(3):
                await session.call_tool("add_slide", {"presentation_id": prs_id})
                await session.call_tool("add_text_box", {
                    "presentation_id": prs_id,
                    "slide_index": i,
                    "text": f"Sample Slide {i + 1}\n\nThis is slide {i + 1} content for testing screenshot functionality.",
                    "font_size": 24,
                    "bold": True
                })
            
            print("✅ Added 3 sample slides with content")
            
            # Save the presentation
            temp_dir = tempfile.mkdtemp()
            ppt_file = os.path.join(temp_dir, "test_presentation.pptx")
            await session.call_tool("save_presentation", {
                "presentation_id": prs_id,
                "file_path": ppt_file
            })
            print(f"✅ Saved presentation to: {ppt_file}")
            
            # Test 3: Take screenshots
            print("\n📋 Step 3: Taking screenshots...")
            
            try:
                screenshot_result = await session.call_tool("screenshot_slides", {
                    "file_path": ppt_file,
                    "image_format": "PNG",
                    "width": 1920,
                    "height": 1080
                })
                
                result_text = screenshot_result.content[0].text
                print("✅ Screenshots created successfully!")
                print(result_text)
                
                # Parse the result to get file paths
                if "screenshot_paths" in result_text:
                    result_lines = result_text.split('\n')
                    json_part = '\n'.join(result_lines[1:])  # Skip first line
                    result_data = json.loads(json_part)
                    
                    print(f"\n📸 Screenshot Details:")
                    print(f"   • Total slides: {result_data['total_slides']}")
                    print(f"   • Image format: {result_data['image_format']}")
                    print(f"   • Dimensions: {result_data['dimensions']}")
                    print(f"   • Output directory: {result_data['output_directory']}")
                    
                    print(f"\n📁 Screenshot files:")
                    for i, path in enumerate(result_data['screenshot_paths'], 1):
                        if os.path.exists(path):
                            file_size = os.path.getsize(path) / 1024  # KB
                            print(f"   • Slide {i}: {os.path.basename(path)} ({file_size:.1f} KB)")
                        else:
                            print(f"   • Slide {i}: {os.path.basename(path)} (FILE NOT FOUND)")
                
                return True
                
            except Exception as e:
                print(f"❌ Screenshot test failed: {e}")
                return False
            
            finally:
                # Cleanup
                try:
                    if os.path.exists(ppt_file):
                        os.remove(ppt_file)
                    if os.path.exists(temp_dir):
                        os.rmdir(temp_dir)
                except:
                    pass


async def demo_usage_example():
    """Demonstrate typical usage of the screenshot feature"""
    
    print("\n" + "=" * 60)
    print("📖 USAGE EXAMPLE")
    print("=" * 60)
    
    print("""
To use the screenshot feature in your applications:

1. **Basic Usage:**
   ```python
   result = await session.call_tool("screenshot_slides", {
       "file_path": "presentation.pptx"
   })
   ```

2. **Custom Settings:**
   ```python
   result = await session.call_tool("screenshot_slides", {
       "file_path": "presentation.pptx",
       "output_dir": "screenshots/",
       "image_format": "PNG",
       "width": 1920,
       "height": 1080
   })
   ```

3. **For AI Vision Analysis:**
   The screenshots can be used with vision AI models to:
   • Analyze slide layouts and design
   • Extract visual elements and charts
   • Generate slide summaries
   • Quality check presentations
   • Create slide thumbnails

**Requirements:**
• Windows operating system
• Microsoft PowerPoint installed
• pywin32 package (pip install pywin32)

**Notes:**
• PowerPoint will briefly open during screenshot process
• Screenshots are saved as high-quality images
• Temporary files are automatically cleaned up
• Works with all PowerPoint formats (.pptx, .ppt)
""")


async def main():
    """Main test function"""
    
    try:
        # Test the functionality
        success = await test_screenshot_functionality()
        
        if success:
            print("\n🎉 All tests passed! Screenshot feature is working correctly.")
            await demo_usage_example()
        else:
            print("\n❌ Tests failed. Please check the error messages above.")
            
    except Exception as e:
        print(f"\n❌ Test execution failed: {e}")
        print(f"\nPossible causes:")
        print("• PowerPoint is not installed")
        print("• pywin32 package is not installed")
        print("• MCP server is not running properly")
        print("• File permissions issue")


if __name__ == "__main__":
    asyncio.run(main()) 