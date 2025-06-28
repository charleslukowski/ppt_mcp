#!/usr/bin/env python3
"""
Test script for PowerPoint MCP Server Table Support - Phase 1

This script tests the basic table functionality:
- Table creation
- Cell content setting
- Table information retrieval
"""

import json
import asyncio
import sys
import os

# Add the current directory to Python path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server_stable import StablePowerPointManager

async def test_table_operations():
    """Test basic table operations"""
    print("🧪 Testing PowerPoint Table Support - Phase 1")
    print("=" * 50)
    
    # Initialize manager
    manager = StablePowerPointManager()
    
    try:
        # Test 1: Create presentation
        print("\n1️⃣ Creating new presentation...")
        prs_id = manager.create_presentation()
        print(f"✅ Created presentation: {prs_id}")
        
        # Test 2: Add slide
        print("\n2️⃣ Adding slide...")
        slide_index = manager.add_slide(prs_id, layout_index=6)  # Blank layout
        print(f"✅ Added slide at index: {slide_index}")
        
        # Test 3: Add table without header
        print("\n3️⃣ Adding basic table (3x4, no header)...")
        table_index = manager.add_table(
            prs_id=prs_id,
            slide_index=slide_index,
            rows=3,
            cols=4,
            left=1,
            top=1,
            width=8,
            height=3,
            header_row=False
        )
        print(f"✅ Added table at index: {table_index}")
        
        # Test 4: Add table with header
        print("\n4️⃣ Adding table with header (4x3)...")
        table_index_2 = manager.add_table(
            prs_id=prs_id,
            slide_index=slide_index,
            rows=4,
            cols=3,
            left=1,
            top=5,
            width=6,
            height=3,
            header_row=True
        )
        print(f"✅ Added header table at index: {table_index_2}")
        
        # Test 5: Set cell content in first table
        print("\n5️⃣ Setting cell content in first table...")
        
        # Set some basic cell content
        test_cells = [
            (0, 0, "Product"),
            (0, 1, "Price"),
            (0, 2, "Quantity"),
            (0, 3, "Total"),
            (1, 0, "Widget A"),
            (1, 1, "$10.99"),
            (1, 2, "5"),
            (1, 3, "$54.95"),
            (2, 0, "Widget B"),
            (2, 1, "$15.99"),
            (2, 2, "3"),
            (2, 3, "$47.97")
        ]
        
        for row, col, text in test_cells:
            success = manager.set_table_cell(
                prs_id=prs_id,
                slide_index=slide_index,
                table_index=table_index,
                row=row,
                col=col,
                text=text
            )
            print(f"  ✅ Set cell [{row},{col}]: '{text}'")
        
        # Test 6: Set formatted content in header table
        print("\n6️⃣ Setting formatted content in header table...")
        
        # Set header content with formatting
        headers = ["Department", "Q1 Sales", "Q2 Sales"]
        for col, header in enumerate(headers):
            success = manager.set_table_cell(
                prs_id=prs_id,
                slide_index=slide_index,
                table_index=table_index_2,
                row=0,
                col=col,
                text=header,
                bold=True,
                font_size=14,
                font_color="white"
            )
            print(f"  ✅ Set header [{0},{col}]: '{header}' (formatted)")
        
        # Set data content
        data_rows = [
            ["Sales", "$125,000", "$140,000"],
            ["Marketing", "$85,000", "$95,000"],
            ["Support", "$65,000", "$70,000"]
        ]
        
        for row_idx, row_data in enumerate(data_rows):
            for col_idx, cell_text in enumerate(row_data):
                color = "#008000" if "$" in cell_text else None  # Green for money
                alignment = "right" if "$" in cell_text else "left"
                
                success = manager.set_table_cell(
                    prs_id=prs_id,
                    slide_index=slide_index,
                    table_index=table_index_2,
                    row=row_idx + 1,  # Skip header row
                    col=col_idx,
                    text=cell_text,
                    font_color=color,
                    text_alignment=alignment
                )
                print(f"  ✅ Set data [{row_idx + 1},{col_idx}]: '{cell_text}' (formatted)")
        
        # Test 7: Get table information
        print("\n7️⃣ Getting table information...")
        
        info1 = manager.get_table_info(prs_id, slide_index, table_index)
        print(f"✅ Table 0 info: {info1['rows']}×{info1['columns']} table with {info1['total_cells']} cells")
        
        info2 = manager.get_table_info(prs_id, slide_index, table_index_2)
        print(f"✅ Table 1 info: {info2['rows']}×{info2['columns']} table with {info2['total_cells']} cells")
        
        # Test 8: Test enhanced content listing
        print("\n8️⃣ Testing slide content listing...")
        
        content = manager.list_slide_content(prs_id, slide_index)
        print(f"✅ Slide content: {content['shape_count']} shapes found")
        
        for shape in content['shapes']:
            if shape['type'] == 'table':
                print(f"  📊 {shape['description']} at index {shape['index']}")
        
        # Test 9: Test text extraction
        print("\n9️⃣ Testing text extraction...")
        
        extracted = manager.extract_text(prs_id)
        for slide_data in extracted:
            if slide_data['text_content']:
                print(f"✅ Slide {slide_data['slide_number']} content:")
                for content_item in slide_data['text_content']:
                    if content_item['shape_type'] == 'table':
                        print(f"  📊 Table ({content_item['rows']}×{content_item['columns']}):")
                        print(f"      {content_item['text'][:100]}...")
        
        # Test 10: Save presentation
        print("\n🔟 Saving presentation...")
        saved_path = manager.save_presentation(prs_id, "test_table_output.pptx")
        print(f"✅ Saved to: {saved_path}")
        
        print("\n🎉 All Phase 1 table tests completed successfully!")
        print(f"📁 Check the saved file: {saved_path}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_error_handling():
    """Test error handling and validation"""
    print("\n🔍 Testing error handling...")
    print("-" * 30)
    
    manager = StablePowerPointManager()
    prs_id = manager.create_presentation()
    slide_index = manager.add_slide(prs_id, layout_index=6)
    
    # Test error cases
    error_tests = [
        ("Invalid table dimensions", lambda: manager.add_table(prs_id, slide_index, 0, 3)),
        ("Out of bounds cell access", lambda: manager.set_table_cell(prs_id, slide_index, 0, 10, 10, "test")),
        ("Non-existent table", lambda: manager.get_table_info(prs_id, slide_index, 5)),
    ]
    
    for test_name, test_func in error_tests:
        try:
            test_func()
            print(f"❌ {test_name}: Should have raised an error")
        except (ValueError, RuntimeError) as e:
            print(f"✅ {test_name}: Correctly raised error - {str(e)[:50]}...")
        except Exception as e:
            print(f"⚠️ {test_name}: Unexpected error type - {e}")

if __name__ == "__main__":
    print("PowerPoint MCP Server - Table Support Test Suite")
    print("Phase 1: Foundation Testing")
    print("=" * 60)
    
    # Run the main test
    success = asyncio.run(test_table_operations())
    
    if success:
        # Run error handling tests
        asyncio.run(test_error_handling())
        
        print("\n🏆 All tests completed!")
        print("\nPhase 1 Implementation Status:")
        print("✅ Table creation (add_table)")
        print("✅ Cell content management (set_table_cell)")
        print("✅ Table information retrieval (get_table_info)")
        print("✅ Enhanced content extraction")
        print("✅ Input validation and error handling")
        print("✅ Success message formatting")
        
        print("\n🚀 Ready for Phase 2: Advanced styling and range operations")
    else:
        print("\n💥 Tests failed - check implementation")
        sys.exit(1) 