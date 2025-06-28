#!/usr/bin/env python3
"""
Test script to simulate the AppData scenario and verify fallback logic

This test simulates running from a problematic directory (like AppData)
to ensure presentations are saved to the Documents folder instead.
"""

import os
import sys
import tempfile
import shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server_stable import StablePowerPointManager

def test_appdata_scenario():
    """Test the scenario where we're running from a problematic directory"""
    print("🚨 Testing AppData/Cursor directory scenario...")
    
    # Create a temporary directory that simulates AppData
    temp_dir = tempfile.mkdtemp()
    appdata_sim = os.path.join(temp_dir, "AppData", "Local", "Programs", "cursor")
    os.makedirs(appdata_sim, exist_ok=True)
    
    # Save current directory
    original_cwd = os.getcwd()
    
    try:
        # Change to the simulated AppData directory
        os.chdir(appdata_sim)
        print(f"Simulated problematic directory: {os.getcwd()}")
        
        manager = StablePowerPointManager()
        
        # Create a presentation
        prs_id = manager.create_presentation()
        slide_index = manager.add_slide(prs_id, 6)
        manager.add_text_box(
            prs_id, slide_index,
            text="AppData Scenario Test",
            font_size=24, text_alignment="center",
            font_color="#FF0000"
        )
        
        # Try to save with relative path - should go to Documents
        relative_file = "appdata_scenario_test.pptx"
        saved_path = manager.save_presentation(prs_id, relative_file)
        
        print(f"✅ Saved to: {saved_path}")
        
        # Verify it went to Documents, not AppData (check directory, not filename)
        saved_dir = os.path.dirname(saved_path)
        if "Documents" in saved_dir and "AppData" not in saved_dir:
            print("✅ Correctly redirected to Documents folder!")
        elif "AppData" in saved_dir:
            print("❌ File saved to AppData (this shouldn't happen)")
        else:
            print("⚠️ File saved to unexpected location")
            
        # Check if file exists
        if os.path.exists(saved_path):
            file_size = os.path.getsize(saved_path)
            print(f"✅ File verified: {file_size} bytes")
        else:
            print("❌ File was not created")
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
    finally:
        # Restore original directory
        os.chdir(original_cwd)
        # Clean up temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)

def test_cursor_directory():
    """Test another problematic directory scenario"""
    print("\n🚨 Testing cursor application directory scenario...")
    
    # Create a temporary directory that simulates cursor app directory
    temp_dir = tempfile.mkdtemp()
    cursor_sim = os.path.join(temp_dir, "cursor", "app")
    os.makedirs(cursor_sim, exist_ok=True)
    
    # Save current directory
    original_cwd = os.getcwd()
    
    try:
        # Change to the simulated cursor directory
        os.chdir(cursor_sim)
        print(f"Simulated cursor directory: {os.getcwd()}")
        
        manager = StablePowerPointManager()
        
        # Create a presentation
        prs_id = manager.create_presentation()
        slide_index = manager.add_slide(prs_id, 6)
        manager.add_text_box(
            prs_id, slide_index,
            text="Cursor Directory Test",
            font_size=24, text_alignment="center",
            font_color="#0066CC"
        )
        
        # Try to save with relative path - should go to Documents
        relative_file = "cursor_scenario_test.pptx"
        saved_path = manager.save_presentation(prs_id, relative_file)
        
        print(f"✅ Saved to: {saved_path}")
        
        # Verify it went to Documents, not cursor directory (check directory, not filename)
        saved_dir = os.path.dirname(saved_path)
        if "Documents" in saved_dir and "cursor" not in saved_dir:
            print("✅ Correctly redirected to Documents folder!")
        elif "cursor" in saved_dir:
            print("❌ File saved to cursor directory (this shouldn't happen)")
        else:
            print("⚠️ File saved to unexpected location")
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
    finally:
        # Restore original directory
        os.chdir(original_cwd)
        # Clean up temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    """Run problematic directory tests"""
    print("🚀 Testing PowerPoint MCP Server problematic directory handling...")
    print("=" * 75)
    
    test_appdata_scenario()
    test_cursor_directory()
    
    print("\n" + "=" * 75)
    print("✅ Problematic directory tests completed!")
    print("\n🎯 These tests verify that presentations are saved to accessible")
    print("   locations even when running from system/application directories.")

if __name__ == "__main__":
    main() 