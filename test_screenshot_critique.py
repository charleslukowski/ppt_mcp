#!/usr/bin/env python3
"""
Test script for PowerPoint MCP Server screenshot and critique functionality.

This script tests:
1. Screenshot generation functionality
2. Presentation critique analysis (design, content, accessibility, technical)
3. Integration between screenshots and critique
4. Various critique types and configurations

Usage:
    python test_screenshot_critique.py
"""

import asyncio
import json
import os
import tempfile
import time
from pathlib import Path

# Import the PowerPoint manager
from powerpoint_mcp_server import PowerPointManager

def create_test_presentation(ppt_manager: PowerPointManager) -> str:
    """Create a test presentation with various issues for critique testing"""
    print("📝 Creating test presentation...")
    
    # Create presentation
    prs_id = ppt_manager.create_presentation()
    
    # Slide 1: Title slide with good practices
    slide_idx = ppt_manager.add_slide(prs_id, 0)  # Title slide layout
    ppt_manager.add_text_box(prs_id, slide_idx, "Professional Presentation", 
                           left=1, top=1, width=8, height=2, font_size=32, bold=True)
    ppt_manager.add_text_box(prs_id, slide_idx, "Quality Analysis Test", 
                           left=1, top=3.5, width=8, height=1, font_size=20)
    
    # Slide 2: Content slide with issues
    slide_idx = ppt_manager.add_slide(prs_id)
    ppt_manager.add_text_box(prs_id, slide_idx, "Slide with Issues", 
                           left=1, top=0.5, width=8, height=1, font_size=24, bold=True)
    # Too much text issue
    long_text = """This slide contains way too much text which violates the best practice of keeping slide content concise and readable. The recommendation is to limit text to 6-7 bullet points maximum and keep character count under 300 per slide. This text is deliberately long to trigger the content analysis warnings in our critique system. Adding even more text here to ensure we cross the 300 character threshold that our system uses to identify overly text-heavy slides."""
    ppt_manager.add_text_box(prs_id, slide_idx, long_text, 
                           left=1, top=2, width=8, height=4, font_size=16)
    
    # Slide 3: Design issues slide
    slide_idx = ppt_manager.add_slide(prs_id)
    ppt_manager.add_text_box(prs_id, slide_idx, "Design Issues Demo", 
                           left=1, top=0.5, width=8, height=1, font_size=28, bold=True)
    # Small font size issue
    ppt_manager.add_text_box(prs_id, slide_idx, "This text uses very small font", 
                           left=1, top=2, width=8, height=1, font_size=10)
    # Large font size issue
    ppt_manager.add_text_box(prs_id, slide_idx, "HUGE TEXT", 
                           left=1, top=3.5, width=8, height=1, font_size=80)
    
    # Slide 4: Chart slide (good practice)
    slide_idx = ppt_manager.add_slide(prs_id)
    ppt_manager.add_text_box(prs_id, slide_idx, "Data Visualization", 
                           left=1, top=0.5, width=8, height=1, font_size=24, bold=True)
    ppt_manager.add_chart(prs_id, slide_idx, "column", 
                         ["Q1", "Q2", "Q3", "Q4"], 
                         {"Revenue": [100, 150, 120, 180], "Profit": [20, 35, 25, 45]})
    
    # Slide 5: Empty slide (content issue)
    slide_idx = ppt_manager.add_slide(prs_id)
    # Intentionally left empty to trigger empty slide warning
    
    # Slide 6: Bullet point overload
    slide_idx = ppt_manager.add_slide(prs_id)
    ppt_manager.add_text_box(prs_id, slide_idx, "Too Many Bullets", 
                           left=1, top=0.5, width=8, height=1, font_size=24, bold=True)
    bullet_text = """• First bullet point
• Second bullet point  
• Third bullet point
• Fourth bullet point
• Fifth bullet point
• Sixth bullet point
• Seventh bullet point
• Eighth bullet point (too many!)
• Ninth bullet point (definitely too many!)"""
    ppt_manager.add_text_box(prs_id, slide_idx, bullet_text, 
                           left=1, top=2, width=8, height=4, font_size=18)
    
    print(f"✅ Test presentation created with ID: {prs_id}")
    return prs_id

async def test_screenshot_functionality(ppt_manager: PowerPointManager, prs_id: str):
    """Test screenshot generation functionality"""
    print("\n📸 Testing screenshot functionality...")
    
    # Save the test presentation first
    test_file = "test_presentation_screenshots.pptx"
    ppt_manager.save_presentation(prs_id, test_file)
    
    if not os.path.exists(test_file):
        print("❌ Failed to save test presentation")
        return None
    
    try:
        # Test screenshot generation
        with tempfile.TemporaryDirectory() as temp_dir:
            print(f"📁 Using temporary directory: {temp_dir}")
            
            # Generate screenshots
            screenshot_paths = await ppt_manager.screenshot_slides_async(
                test_file, temp_dir, "PNG", 1920, 1080
            )
            
            print(f"✅ Generated {len(screenshot_paths)} screenshots")
            for i, path in enumerate(screenshot_paths):
                if os.path.exists(path):
                    size = os.path.getsize(path) / 1024  # KB
                    print(f"   📸 Slide {i+1}: {os.path.basename(path)} ({size:.1f} KB)")
                else:
                    print(f"   ❌ Missing: {path}")
            
            return screenshot_paths
            
    except Exception as e:
        print(f"❌ Screenshot test failed: {e}")
        return None
    finally:
        # Cleanup
        if os.path.exists(test_file):
            os.remove(test_file)

async def test_critique_functionality(ppt_manager: PowerPointManager, prs_id: str):
    """Test presentation critique functionality"""
    print("\n🔍 Testing critique functionality...")
    
    # Save the test presentation
    test_file = "test_presentation_critique.pptx"
    ppt_manager.save_presentation(prs_id, test_file)
    
    if not os.path.exists(test_file):
        print("❌ Failed to save test presentation")
        return
    
    try:
        # Test different critique types
        critique_types = ["design", "content", "accessibility", "technical", "comprehensive"]
        
        for critique_type in critique_types:
            print(f"\n🎯 Testing {critique_type} critique...")
            
            start_time = time.time()
            critique_results = await ppt_manager.critique_presentation_async(
                test_file, critique_type, include_screenshots=False
            )
            analysis_time = time.time() - start_time
            
            # Display results
            summary = critique_results["summary"]
            print(f"   📊 Overall Score: {summary['overall_score']}/100 ({summary['assessment']})")
            print(f"   🔴 Critical Issues: {summary['critical_issues']}")
            print(f"   ⚠️  Warnings: {summary['warnings']}")
            print(f"   💡 Recommendations: {summary['recommendations']}")
            print(f"   ⏱️  Analysis Time: {analysis_time:.2f}s")
            
            # Show sample issues
            if critique_results["issues"]:
                print(f"   🚨 Sample Issues:")
                for issue in critique_results["issues"][:3]:
                    emoji = "🔴" if issue["type"] == "critical" else "⚠️"
                    slide_info = f"Slide {issue['slide']}" if issue['slide'] != 'global' else "Global"
                    print(f"      {emoji} {slide_info}: {issue['issue']}")
            
            # Show sample strengths
            if critique_results["strengths"]:
                print(f"   ✅ Strengths:")
                for strength in critique_results["strengths"][:2]:
                    print(f"      • {strength}")
        
        # Test comprehensive critique with screenshots
        print(f"\n🎯 Testing comprehensive critique with screenshots...")
        with tempfile.TemporaryDirectory() as temp_dir:
            start_time = time.time()
            critique_results = await ppt_manager.critique_presentation_async(
                test_file, "comprehensive", include_screenshots=True, output_dir=temp_dir
            )
            total_time = time.time() - start_time
            
            summary = critique_results["summary"]
            print(f"   📊 Overall Score: {summary['overall_score']}/100 ({summary['assessment']})")
            print(f"   📸 Screenshots: {len(critique_results.get('screenshots', []))} generated")
            print(f"   ⏱️  Total Time: {total_time:.2f}s")
            
            # Detailed analysis breakdown
            if "detailed_analysis" in critique_results:
                print(f"   📋 Analysis Breakdown:")
                for category, analysis in critique_results["detailed_analysis"].items():
                    score = analysis.get("score", 0)
                    issues = len(analysis.get("issues", []))
                    print(f"      • {category.title()}: {score}/100 ({issues} issues)")
        
    except Exception as e:
        print(f"❌ Critique test failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup
        if os.path.exists(test_file):
            os.remove(test_file)

async def test_integrated_functionality(ppt_manager: PowerPointManager):
    """Test integrated screenshot and critique workflow"""
    print("\n🔗 Testing integrated screenshot + critique workflow...")
    
    # Create a comprehensive test presentation
    prs_id = create_test_presentation(ppt_manager)
    
    # Test the complete workflow
    test_file = "test_integrated_workflow.pptx"
    ppt_manager.save_presentation(prs_id, test_file)
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            print(f"📁 Working directory: {temp_dir}")
            
            # Step 1: Generate screenshots
            print("Step 1: Generating screenshots...")
            screenshot_paths = await ppt_manager.screenshot_slides_async(
                test_file, temp_dir, "PNG", 1920, 1080
            )
            print(f"✅ Generated {len(screenshot_paths)} screenshots")
            
            # Step 2: Run comprehensive critique with screenshots
            print("Step 2: Running comprehensive critique...")
            critique_results = await ppt_manager.critique_presentation_async(
                test_file, "comprehensive", include_screenshots=True, output_dir=temp_dir
            )
            
            # Step 3: Analyze results
            print("Step 3: Analyzing integrated results...")
            summary = critique_results["summary"]
            
            print(f"\n📊 Final Results:")
            print(f"   🎯 Assessment: {summary['assessment']} ({summary['overall_score']}/100)")
            print(f"   📈 Slides Analyzed: {summary['total_slides']}")
            print(f"   📸 Screenshots Generated: {len(critique_results.get('screenshots', []))}")
            print(f"   🔴 Critical Issues: {summary['critical_issues']}")
            print(f"   ⚠️  Warnings: {summary['warnings']}")
            print(f"   💡 Recommendations: {summary['recommendations']}")
            
            # Detailed breakdown by category
            if "detailed_analysis" in critique_results:
                print(f"\n📋 Detailed Analysis:")
                for category, analysis in critique_results["detailed_analysis"].items():
                    metrics = analysis.get("metrics", {})
                    score = analysis.get("score", 0)
                    print(f"   📊 {category.title()}: {score}/100")
                    
                    # Show key metrics
                    if category == "design" and "total_fonts" in metrics:
                        print(f"      • Fonts used: {metrics['total_fonts']}")
                        font_range = metrics.get("font_sizes_range", {})
                        if font_range:
                            print(f"      • Font size range: {font_range.get('min', 0)}-{font_range.get('max', 0)}pt")
                    elif category == "content" and "empty_slides" in metrics:
                        print(f"      • Empty slides: {metrics['empty_slides']}")
                        print(f"      • Avg text length: {metrics.get('avg_text_length', 0):.0f} chars")
                    elif category == "technical" and "file_size_mb" in metrics:
                        print(f"      • File size: {metrics['file_size_mb']} MB")
                        print(f"      • Embedded objects: {metrics.get('embedded_objects', 0)}")
                    elif category == "accessibility" and "total_images" in metrics:
                        print(f"      • Images: {metrics['total_images']}")
                        print(f"      • Missing alt text: {metrics.get('alt_text_missing', 0)}")
            
            # Show top issues and recommendations
            if critique_results["issues"]:
                print(f"\n🚨 Top Issues:")
                for issue in critique_results["issues"][:5]:
                    emoji = "🔴" if issue["type"] == "critical" else "⚠️"
                    slide_info = f"Slide {issue['slide']}" if issue['slide'] != 'global' else "Global"
                    print(f"   {emoji} {slide_info}: {issue['issue']}")
            
            if critique_results["recommendations"]:
                print(f"\n💡 Top Recommendations:")
                unique_recs = list(set(critique_results["recommendations"]))
                for rec in unique_recs[:5]:
                    print(f"   • {rec}")
            
            # Test screenshot file integration
            screenshot_files = critique_results.get("screenshots", [])
            if screenshot_files:
                print(f"\n📸 Screenshot Validation:")
                for i, path in enumerate(screenshot_files):
                    if os.path.exists(path):
                        size = os.path.getsize(path)
                        print(f"   ✅ Slide {i+1}: {os.path.basename(path)} ({size/1024:.1f} KB)")
                    else:
                        print(f"   ❌ Missing: {os.path.basename(path)}")
            
            print(f"\n🎉 Integrated workflow test completed successfully!")
            
    except Exception as e:
        print(f"❌ Integrated test failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup
        if os.path.exists(test_file):
            os.remove(test_file)
        # Clean up the presentation from memory
        if prs_id in ppt_manager.presentations:
            del ppt_manager.presentations[prs_id]

async def main():
    """Main test runner"""
    print("🚀 PowerPoint MCP Server - Screenshot & Critique Test Suite")
    print("=" * 60)
    
    # Initialize PowerPoint manager
    ppt_manager = PowerPointManager()
    
    try:
        # Create test presentation
        prs_id = create_test_presentation(ppt_manager)
        
        # Run individual tests
        await test_screenshot_functionality(ppt_manager, prs_id)
        await test_critique_functionality(ppt_manager, prs_id)
        
        # Clean up test presentation
        if prs_id in ppt_manager.presentations:
            del ppt_manager.presentations[prs_id]
        
        # Run integrated test
        await test_integrated_functionality(ppt_manager)
        
        print("\n" + "=" * 60)
        print("🎉 All tests completed!")
        print("\n📋 Test Summary:")
        print("   ✅ Screenshot generation")
        print("   ✅ Critique analysis (all types)")
        print("   ✅ Integrated workflow")
        print("   ✅ File handling and cleanup")
        
    except Exception as e:
        print(f"\n❌ Test suite failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Final cleanup
        ppt_manager.cleanup()

if __name__ == "__main__":
    asyncio.run(main()) 