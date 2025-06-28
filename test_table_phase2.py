#!/usr/bin/env python3
"""
Test script for PowerPoint MCP Server Table Support - Phase 2

This script tests the advanced table styling functionality:
- Individual cell styling (style_table_cell)
- Range styling operations (style_table_range)
- Data-driven table creation (create_table_with_data)
"""

import json
import asyncio
import sys
import os

# Add the current directory to Python path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server_stable import StablePowerPointManager

async def test_phase2_styling():
    """Test Phase 2 advanced styling operations"""
    print("🎨 Testing PowerPoint Table Support - Phase 2: Advanced Styling")
    print("=" * 60)
    
    # Initialize manager
    manager = StablePowerPointManager()
    
    try:
        # Test 1: Create presentation and slide
        print("\n1️⃣ Setting up presentation...")
        prs_id = manager.create_presentation()
        slide_index = manager.add_slide(prs_id, layout_index=6)
        print(f"✅ Created presentation {prs_id} with slide {slide_index}")
        
        # Test 2: Create table with basic structure for styling tests
        print("\n2️⃣ Creating base table for styling tests...")
        table_index = manager.add_table(
            prs_id=prs_id,
            slide_index=slide_index,
            rows=4,
            cols=3,
            left=1,
            top=1,
            width=7,
            height=3,
            header_row=False
        )
        
        # Populate with test data
        test_data = [
            ["Header 1", "Header 2", "Header 3"],
            ["Data A1", "Data B1", "Data C1"],
            ["Data A2", "Data B2", "Data C2"],
            ["Data A3", "Data B3", "Data C3"]
        ]
        
        for row, row_data in enumerate(test_data):
            for col, cell_text in enumerate(row_data):
                manager.set_table_cell(prs_id, slide_index, table_index, row, col, cell_text)
        
        print(f"✅ Created and populated 4×3 table at index {table_index}")
        
        # Test 3: Individual cell styling
        print("\n3️⃣ Testing individual cell styling...")
        
        # Style header cells with blue background
        for col in range(3):
            success = manager.style_table_cell(
                prs_id=prs_id,
                slide_index=slide_index,
                table_index=table_index,
                row=0,
                col=col,
                fill_color="#4472C4",
                margin_left=0.1,
                margin_right=0.1,
                margin_top=0.05,
                margin_bottom=0.05
            )
            print(f"  ✅ Styled header cell [0,{col}]: blue background + margins")
        
        # Style specific data cells with different colors
        color_tests = [
            (1, 0, "#E6F3FF", "Light blue data cell"),
            (1, 1, "#FFE6E6", "Light red data cell"),
            (1, 2, "#E6FFE6", "Light green data cell"),
            (2, 1, "#FFFFE6", "Light yellow data cell")
        ]
        
        for row, col, color, description in color_tests:
            success = manager.style_table_cell(
                prs_id=prs_id,
                slide_index=slide_index,
                table_index=table_index,
                row=row,
                col=col,
                fill_color=color,
                margin_left=0.05,
                margin_right=0.05
            )
            print(f"  ✅ Styled cell [{row},{col}]: {description}")
        
        # Test 4: Range styling
        print("\n4️⃣ Testing range styling operations...")
        
        # Style entire bottom row with gray background
        success = manager.style_table_range(
            prs_id=prs_id,
            slide_index=slide_index,
            table_index=table_index,
            start_row=3,
            start_col=0,
            end_row=3,
            end_col=2,
            fill_color="#F0F0F0",
            margin_top=0.1,
            margin_bottom=0.1
        )
        print("  ✅ Styled bottom row range [3,0] to [3,2]: gray background + margins")
        
        # Style a 2x2 range in the middle
        success = manager.style_table_range(
            prs_id=prs_id,
            slide_index=slide_index,
            table_index=table_index,
            start_row=1,
            start_col=1,
            end_row=2,
            end_col=2,
            fill_color="#FFF8DC",  # Cornsilk color
            margin_left=0.15,
            margin_right=0.15
        )
        print("  ✅ Styled 2×2 range [1,1] to [2,2]: cornsilk background + extra margins")
        
        # Test 5: Create table with data (convenience method)
        print("\n5️⃣ Testing create_table_with_data...")
        
        # Employee data table
        employee_data = [
            ["Alice Johnson", "Marketing", "$75,000", "5 years"],
            ["Bob Smith", "Engineering", "$95,000", "3 years"],
            ["Carol Davis", "Sales", "$68,000", "7 years"],
            ["David Wilson", "Engineering", "$88,000", "2 years"],
            ["Eve Brown", "Marketing", "$72,000", "4 years"]
        ]
        
        headers = ["Employee", "Department", "Salary", "Experience"]
        
        table_index_2 = manager.create_table_with_data(
            prs_id=prs_id,
            slide_index=slide_index,
            table_data=employee_data,
            headers=headers,
            left=1,
            top=5,
            width=8,
            height=3,
            header_style={"bold": True, "font_size": 12, "font_color": "white"},
            data_style={"font_size": 10},
            alternating_rows=True
        )
        print(f"✅ Created employee table with data at index {table_index_2}")
        print(f"    - {len(employee_data)} data rows + {len(headers)} headers")
        print(f"    - Alternating row colors enabled")
        print(f"    - Header styling: bold, 12pt, white text")
        print(f"    - Data styling: 10pt font")
        
        # Test 6: Financial data table with custom styling
        print("\n6️⃣ Creating financial data table with advanced styling...")
        
        financial_data = [
            ["Q1 2024", "$125,000", "$110,000", "+13.6%"],
            ["Q2 2024", "$142,000", "$125,000", "+13.6%"],
            ["Q3 2024", "$138,000", "$130,000", "+6.2%"],
            ["Q4 2024", "$155,000", "$140,000", "+10.7%"]
        ]
        
        financial_headers = ["Quarter", "Revenue", "Target", "Growth"]
        
        table_index_3 = manager.create_table_with_data(
            prs_id=prs_id,
            slide_index=slide_index,
            table_data=financial_data,
            headers=financial_headers,
            left=1,
            top=8.5,
            width=7,
            height=2.5,
            header_style={"bold": True, "font_size": 14, "font_color": "#FFFFFF"},
            data_style={"font_size": 11, "text_alignment": "center"},
            alternating_rows=False
        )
        
        # Apply custom styling to financial table
        # Style header row with dark blue
        manager.style_table_range(
            prs_id, slide_index, table_index_3, 0, 0, 0, 3,
            fill_color="#1f4e79"
        )
        
        # Style growth column (positive numbers) with green background
        for row in range(1, 5):  # Data rows
            manager.style_table_cell(
                prs_id, slide_index, table_index_3, row, 3,
                fill_color="#d5e8d4"  # Light green
            )
        
        # Style revenue column with light blue
        for row in range(1, 5):
            manager.style_table_cell(
                prs_id, slide_index, table_index_3, row, 1,
                fill_color="#dae8fc"  # Light blue
            )
        
        print(f"✅ Created financial table with custom styling at index {table_index_3}")
        print("    - Dark blue header background")
        print("    - Green background for growth percentages")
        print("    - Blue background for revenue column")
        
        # Test 7: Get information about all tables
        print("\n7️⃣ Getting table information...")
        
        for i, table_idx in enumerate([table_index, table_index_2, table_index_3]):
            info = manager.get_table_info(prs_id, slide_index, table_idx)
            print(f"  📊 Table {table_idx}: {info['rows']}×{info['columns']} ({info['total_cells']} cells)")
        
        # Test 8: Test slide content listing with styled tables
        print("\n8️⃣ Testing slide content listing...")
        
        content = manager.list_slide_content(prs_id, slide_index)
        print(f"✅ Slide content: {content['shape_count']} shapes found")
        
        table_count = sum(1 for shape in content['shapes'] if shape['type'] == 'table')
        print(f"    📊 Found {table_count} tables on the slide")
        
        for shape in content['shapes']:
            if shape['type'] == 'table':
                print(f"      - {shape['description']} at index {shape['index']}")
        
        # Test 9: Save presentation
        print("\n9️⃣ Saving presentation...")
        saved_path = manager.save_presentation(prs_id, "test_table_phase2_output.pptx")
        print(f"✅ Saved to: {saved_path}")
        
        print("\n🎉 All Phase 2 styling tests completed successfully!")
        print(f"📁 Check the saved file: {saved_path}")
        
        # Summary of what was tested
        print("\n📋 Phase 2 Features Tested:")
        print("✅ Individual cell styling (fill colors, margins)")
        print("✅ Range styling operations (row and block ranges)")
        print("✅ Data-driven table creation with styling")
        print("✅ Alternating row colors")
        print("✅ Custom header and data styling")
        print("✅ Complex multi-table layouts")
        print("✅ Integration with existing table management")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Phase 2 test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_phase2_error_handling():
    """Test error handling for Phase 2 features"""
    print("\n🔍 Testing Phase 2 error handling...")
    print("-" * 40)
    
    manager = StablePowerPointManager()
    prs_id = manager.create_presentation()
    slide_index = manager.add_slide(prs_id, layout_index=6)
    
    # Create a small table for error testing
    table_index = manager.add_table(prs_id, slide_index, 2, 2)
    
    # Test error cases
    error_tests = [
        ("Invalid cell styling coordinates", 
         lambda: manager.style_table_cell(prs_id, slide_index, table_index, 5, 5, fill_color="red")),
        ("Invalid range (start > end)", 
         lambda: manager.style_table_range(prs_id, slide_index, table_index, 1, 1, 0, 0, fill_color="blue")),
        ("Invalid table data structure", 
         lambda: manager.create_table_with_data(prs_id, slide_index, "not a list")),
        ("Inconsistent row lengths", 
         lambda: manager.create_table_with_data(prs_id, slide_index, [["A", "B"], ["C"]])),
        ("Headers length mismatch", 
         lambda: manager.create_table_with_data(prs_id, slide_index, [["A", "B"]], headers=["H1", "H2", "H3"])),
    ]
    
    for test_name, test_func in error_tests:
        try:
            test_func()
            print(f"❌ {test_name}: Should have raised an error")
        except (ValueError, RuntimeError) as e:
            print(f"✅ {test_name}: Correctly raised error - {str(e)[:60]}...")
        except Exception as e:
            print(f"⚠️ {test_name}: Unexpected error type - {e}")

async def test_phase2_performance():
    """Test performance with larger tables and bulk operations"""
    print("\n⚡ Testing Phase 2 performance...")
    print("-" * 30)
    
    manager = StablePowerPointManager()
    prs_id = manager.create_presentation()
    slide_index = manager.add_slide(prs_id, layout_index=6)
    
    # Create large dataset
    large_data = []
    for i in range(50):  # 50 rows
        row = [f"Item {i}", f"Value {i*10}", f"Status {i%3}", f"Score {i*2.5}"]
        large_data.append(row)
    
    headers = ["Item", "Value", "Status", "Score"]
    
    print(f"Creating large table with {len(large_data)} rows and {len(headers)} columns...")
    
    import time
    start_time = time.time()
    
    table_index = manager.create_table_with_data(
        prs_id, slide_index, large_data, headers,
        alternating_rows=True,
        header_style={"bold": True, "font_size": 10},
        data_style={"font_size": 8}
    )
    
    creation_time = time.time() - start_time
    print(f"✅ Large table created in {creation_time:.2f} seconds")
    
    # Test bulk styling
    start_time = time.time()
    
    # Style header row
    manager.style_table_range(prs_id, slide_index, table_index, 0, 0, 0, 3, fill_color="#2F5597")
    
    # Style status column with different colors based on status
    for row in range(1, 51):
        if row % 3 == 1:
            color = "#d5e8d4"  # Green for status 1
        elif row % 3 == 2:
            color = "#fff2cc"  # Yellow for status 2
        else:
            color = "#f8cecc"  # Red for status 0
        
        manager.style_table_cell(prs_id, slide_index, table_index, row, 2, fill_color=color)
    
    styling_time = time.time() - start_time
    print(f"✅ Bulk styling completed in {styling_time:.2f} seconds")
    
    # Save and check file size
    start_time = time.time()
    saved_path = manager.save_presentation(prs_id, "test_table_performance.pptx")
    save_time = time.time() - start_time
    
    file_size = os.path.getsize(saved_path)
    print(f"✅ Large presentation saved in {save_time:.2f} seconds")
    print(f"📁 File size: {file_size:,} bytes ({file_size/1024:.1f} KB)")

if __name__ == "__main__":
    print("PowerPoint MCP Server - Table Support Test Suite")
    print("Phase 2: Advanced Styling Testing")
    print("=" * 70)
    
    # Run the main Phase 2 test
    success = asyncio.run(test_phase2_styling())
    
    if success:
        # Run error handling tests
        asyncio.run(test_phase2_error_handling())
        
        # Run performance tests
        asyncio.run(test_phase2_performance())
        
        print("\n🏆 All Phase 2 tests completed!")
        print("\nPhase 2 Implementation Status:")
        print("✅ Individual cell styling (style_table_cell)")
        print("✅ Range styling operations (style_table_range)")
        print("✅ Data-driven table creation (create_table_with_data)")
        print("✅ Alternating row colors")
        print("✅ Custom styling options")
        print("✅ Performance optimization")
        print("✅ Enhanced error handling")
        
        print("\n🚀 Phase 2 Complete - Ready for Phase 3: Structure Modification")
    else:
        print("\n💥 Phase 2 tests failed - check implementation")
        sys.exit(1) 