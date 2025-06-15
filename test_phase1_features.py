#!/usr/bin/env python3
"""
Test Phase 1: Professional Formatting & Layout Features

This test suite validates all the Phase 1 features implemented in the PowerPoint MCP server:
- Grid-Based Positioning
- Master Slide Management  
- Typography System
- Shape Libraries
- Color Palette Management

Usage:
    python test_phase1_features.py
"""

import os
import sys
import json
import tempfile
from pathlib import Path

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server import PowerPointManager

def test_grid_based_positioning():
    """Test grid-based positioning features"""
    print("🔲 Testing Grid-Based Positioning...")
    
    manager = PowerPointManager()
    
    # Create a presentation
    prs_id = manager.create_presentation()
    slide_index = manager.add_slide(prs_id)
    
    # Test creating layout grid
    success = manager.create_layout_grid(prs_id, columns=4, rows=3)
    assert success, "Failed to create layout grid"
    print("  ✅ Layout grid created successfully")
    
    # Add some shapes to test positioning
    manager.add_text_box(prs_id, slide_index, "Shape 1", left=1, top=1)
    manager.add_text_box(prs_id, slide_index, "Shape 2", left=3, top=1)
    manager.add_text_box(prs_id, slide_index, "Shape 3", left=5, top=1)
    
    # Test snapping to grid
    success = manager.snap_to_grid(prs_id, slide_index, "0", (0, 0))
    assert success, "Failed to snap shape to grid"
    print("  ✅ Shape snapped to grid successfully")
    
    # Test shape distribution
    success = manager.distribute_shapes(prs_id, slide_index, ["0", "1", "2"], "horizontal")
    assert success, "Failed to distribute shapes"
    print("  ✅ Shapes distributed successfully")
    
    print("  🎉 Grid-Based Positioning tests passed!\n")
    return prs_id, manager

def test_color_palette_management(prs_id, manager):
    """Test color palette management features"""
    print("🎨 Testing Color Palette Management...")
    
    # Test creating predefined color palette
    success = manager.create_color_palette(prs_id, "corporate_blue")
    assert success, "Failed to create predefined color palette"
    print("  ✅ Predefined color palette created successfully")
    
    # Test creating custom color palette
    custom_colors = {
        "primary": "#FF5722",
        "secondary": "#FFC107", 
        "accent": "#4CAF50",
        "text_dark": "#212121",
        "text_light": "#FFFFFF"
    }
    success = manager.create_color_palette(prs_id, "custom_palette", custom_colors)
    assert success, "Failed to create custom color palette"
    print("  ✅ Custom color palette created successfully")
    
    # Test applying color palette
    success = manager.apply_color_palette(prs_id, "corporate_blue")
    assert success, "Failed to apply color palette"
    print("  ✅ Color palette applied successfully")
    
    print("  🎉 Color Palette Management tests passed!\n")

def test_typography_system(prs_id, manager):
    """Test typography system features"""
    print("📝 Testing Typography System...")
    slide_index = manager.add_slide(prs_id)
    
    # Create custom typography profile
    typography_config = {
        "title": {"font_name": "Arial", "font_size": 48, "bold": True, "color": "primary"},
        "subtitle": {"font_name": "Arial", "font_size": 28, "bold": False, "color": "secondary"},
        "heading": {"font_name": "Arial", "font_size": 20, "bold": True, "color": "text_dark"},
        "body": {"font_name": "Arial", "font_size": 16, "bold": False, "color": "text_dark"},
        "caption": {"font_name": "Arial", "font_size": 12, "bold": False, "color": "secondary"}
    }
    
    success = manager.create_typography_profile(prs_id, "modern_profile", typography_config)
    assert success, "Failed to create typography profile"
    print("  ✅ Typography profile created successfully")
    
    # Add text shapes to test typography
    manager.add_text_box(prs_id, slide_index, "Main Title", top=0.5, font_size=24)
    manager.add_text_box(prs_id, slide_index, "Body Content", top=2, font_size=18)
    
    # Test applying typography styles
    success = manager.apply_typography_style(prs_id, slide_index, "0", "title", "modern_profile")
    assert success, "Failed to apply title typography"
    print("  ✅ Title typography applied successfully")
    
    success = manager.apply_typography_style(prs_id, slide_index, "1", "body", "modern_profile")
    assert success, "Failed to apply body typography"
    print("  ✅ Body typography applied successfully")
    
    print("  🎉 Typography System tests passed!\n")

def test_shape_libraries(prs_id, manager):
    """Test professional shape libraries"""
    print("🔷 Testing Shape Libraries...")
    slide_index = manager.add_slide(prs_id)
    
    # Test listing shape library
    library = manager.list_shape_library()
    assert len(library) > 0, "Shape library is empty"
    assert "arrows" in library, "Arrows category not found in library"
    assert "geometric" in library, "Geometric category not found in library"
    print("  ✅ Shape library listed successfully")
    print(f"      Available categories: {list(library.keys())}")
    
    # Test adding professional shapes
    success = manager.add_professional_shape(prs_id, slide_index, "arrows", "0", left=1, top=1)
    assert success, "Failed to add arrow shape"
    print("  ✅ Arrow shape added successfully")
    
    success = manager.add_professional_shape(prs_id, slide_index, "geometric", "0", left=3, top=1)
    assert success, "Failed to add geometric shape"
    print("  ✅ Geometric shape added successfully")
    
    success = manager.add_professional_shape(prs_id, slide_index, "callouts", "0", left=5, top=1)
    assert success, "Failed to add callout shape"
    print("  ✅ Callout shape added successfully")
    
    print("  🎉 Shape Libraries tests passed!\n")

def test_master_slide_management(prs_id, manager):
    """Test master slide management features"""
    print("🎭 Testing Master Slide Management...")
    
    # Test creating master slide theme
    theme_config = {
        "background_color": "#F5F5F5",
        "title_font": {
            "name": "Segoe UI",
            "size": 36,
            "color": "#2E3440",
            "bold": True
        },
        "content_font": {
            "name": "Segoe UI",
            "size": 16,
            "color": "#3B4252",
            "bold": False
        },
        "accent_color": "#5E81AC"
    }
    
    success = manager.create_master_slide_theme(prs_id, "modern_theme", theme_config)
    assert success, "Failed to create master slide theme"
    print("  ✅ Master slide theme created successfully")
    
    # Test listing master themes
    themes = manager.list_master_themes(prs_id)
    assert "modern_theme" in themes, "Created theme not found in list"
    print("  ✅ Master themes listed successfully")
    print(f"      Available themes: {themes}")
    
    # Add some content to test theme application
    slide_index = manager.add_slide(prs_id)
    manager.add_text_box(prs_id, slide_index, "Theme Test Title", top=1, font_size=32)
    manager.add_text_box(prs_id, slide_index, "Theme test content", top=3, font_size=18)
    
    # Test applying master theme
    success = manager.apply_master_theme(prs_id, "modern_theme")
    assert success, "Failed to apply master theme"
    print("  ✅ Master theme applied successfully")
    
    # Test slide layout templates
    template_config = {
        "layout_type": "title_content",
        "title": "Template Title",
        "content": "This is content added via template",
        "clear_existing": False
    }
    
    slide_index = manager.add_slide(prs_id)
    success = manager.set_slide_layout_template(prs_id, slide_index, template_config)
    assert success, "Failed to set slide layout template"
    print("  ✅ Slide layout template applied successfully")
    
    # Test two-column layout
    template_config = {
        "layout_type": "two_column",
        "left_content": "Left column content",
        "right_content": "Right column content"
    }
    
    slide_index = manager.add_slide(prs_id)
    success = manager.set_slide_layout_template(prs_id, slide_index, template_config)
    assert success, "Failed to set two-column layout template"
    print("  ✅ Two-column layout template applied successfully")
    
    print("  🎉 Master Slide Management tests passed!\n")

def test_comprehensive_workflow(prs_id, manager):
    """Test a comprehensive workflow using all Phase 1 features"""
    print("🚀 Testing Comprehensive Phase 1 Workflow...")
    
    # Create a professional presentation using all features
    slide_index = manager.add_slide(prs_id)
    
    # 1. Set up grid layout
    manager.create_layout_grid(prs_id, columns=3, rows=2, margins={"left": 1, "right": 1, "top": 1, "bottom": 1})
    
    # 2. Create and apply color palette
    brand_colors = {
        "primary": "#1E3A8A",     # Blue
        "secondary": "#059669",   # Green
        "accent": "#DC2626",      # Red
        "text_dark": "#1F2937",   # Dark gray
        "text_light": "#F9FAFB"   # Light gray
    }
    manager.create_color_palette(prs_id, "brand_palette", brand_colors)
    
    # 3. Create typography profile
    typography_config = {
        "title": {"font_name": "Segoe UI", "font_size": 42, "bold": True, "color": "primary"},
        "subtitle": {"font_name": "Segoe UI", "font_size": 24, "bold": False, "color": "secondary"},
        "body": {"font_name": "Segoe UI", "font_size": 16, "bold": False, "color": "text_dark"}
    }
    manager.create_typography_profile(prs_id, "brand_typography", typography_config)
    
    # 4. Create master theme
    master_theme = {
        "background_color": "#FFFFFF",
        "title_font": {"name": "Segoe UI", "size": 42, "color": "#1E3A8A", "bold": True},
        "content_font": {"name": "Segoe UI", "size": 16, "color": "#1F2937", "bold": False},
    }
    manager.create_master_slide_theme(prs_id, "brand_master", master_theme)
    
    # 5. Add content using layout template
    template_config = {
        "layout_type": "title_content",
        "title": "Phase 1 Features Demo",
        "content": "This presentation demonstrates all Phase 1 professional formatting features"
    }
    manager.set_slide_layout_template(prs_id, slide_index, template_config)
    
    # 6. Add professional shapes
    manager.add_professional_shape(prs_id, slide_index, "arrows", "1", left=7, top=5, width=1.5, height=1)
    manager.add_professional_shape(prs_id, slide_index, "geometric", "2", left=8.5, top=5, width=1, height=1)
    
    # 7. Apply all formatting
    manager.apply_color_palette(prs_id, "brand_palette")
    manager.apply_master_theme(prs_id, "brand_master")
    
    print("  ✅ Comprehensive workflow completed successfully")
    print("  🎉 All Phase 1 features integrated successfully!\n")

def main():
    """Run all Phase 1 feature tests"""
    print("🔬 Testing Phase 1: Professional Formatting & Layout Features")
    print("=" * 70)
    
    try:
        # Test each feature category
        prs_id, manager = test_grid_based_positioning()
        test_color_palette_management(prs_id, manager)
        test_typography_system(prs_id, manager)
        test_shape_libraries(prs_id, manager)
        test_master_slide_management(prs_id, manager)
        test_comprehensive_workflow(prs_id, manager)
        
        # Save the test presentation
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp_file:
            test_file_path = tmp_file.name
        
        manager.save_presentation(prs_id, test_file_path)
        
        print("🎊 ALL PHASE 1 TESTS PASSED! 🎊")
        print("=" * 70)
        print(f"📄 Test presentation saved to: {test_file_path}")
        print("\n🌟 Phase 1 Professional Formatting & Layout features are ready!")
        print("\n📋 Summary of implemented features:")
        print("  ✅ Grid-Based Positioning (create_layout_grid, snap_to_grid, distribute_shapes)")
        print("  ✅ Color Palette Management (create_color_palette, apply_color_palette)")
        print("  ✅ Typography System (create_typography_profile, apply_typography_style)")
        print("  ✅ Shape Libraries (add_professional_shape, list_shape_library)")
        print("  ✅ Master Slide Management (create_master_slide_theme, apply_master_theme)")
        print("  ✅ Layout Templates (set_slide_layout_template)")
        
        # Clean up
        manager.cleanup()
        
        return True
        
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 