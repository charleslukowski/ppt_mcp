#!/usr/bin/env python3
"""
Example usage of PowerPoint MCP Server screenshot and critique features.

This script demonstrates:
1. Creating a sample presentation
2. Taking screenshots
3. Running critique analysis
4. Displaying results

Usage:
    python example_screenshot_critique.py
"""

import asyncio
import json
import os
import tempfile
from powerpoint_mcp_server import PowerPointManager

async def main():
    """Example workflow demonstrating screenshot and critique features"""
    print("🚀 PowerPoint MCP Server - Screenshot & Critique Example")
    print("=" * 55)
    
    # Initialize PowerPoint manager
    ppt_manager = PowerPointManager()
    
    try:
        # Step 1: Create a sample presentation
        print("\n📝 Step 1: Creating sample presentation...")
        prs_id = ppt_manager.create_presentation()
        
        # Add title slide
        slide_idx = ppt_manager.add_slide(prs_id, 0)  # Title layout
        ppt_manager.add_text_box(prs_id, slide_idx, "Sample Presentation", 
                               left=1, top=2, width=8, height=1.5, font_size=32, bold=True)
        ppt_manager.add_text_box(prs_id, slide_idx, "Screenshot & Critique Demo", 
                               left=1, top=4, width=8, height=1, font_size=20)
        
        # Add content slide
        slide_idx = ppt_manager.add_slide(prs_id)
        ppt_manager.add_text_box(prs_id, slide_idx, "Key Features", 
                               left=1, top=1, width=8, height=1, font_size=24, bold=True)
        ppt_manager.add_text_box(prs_id, slide_idx, 
                               "• Screenshot generation\n• Design analysis\n• Content critique\n• Accessibility review", 
                               left=1, top=2.5, width=8, height=3, font_size=18)
        
        # Add chart slide
        slide_idx = ppt_manager.add_slide(prs_id)
        ppt_manager.add_text_box(prs_id, slide_idx, "Performance Metrics", 
                               left=1, top=0.5, width=8, height=1, font_size=24, bold=True)
        ppt_manager.add_chart(prs_id, slide_idx, "column", 
                             ["Speed", "Quality", "Accuracy"], 
                             {"Results": [85, 92, 88]})
        
        print("✅ Sample presentation created with 3 slides")
        
        # Step 2: Save the presentation
        print("\n💾 Step 2: Saving presentation...")
        sample_file = "sample_presentation.pptx"
        ppt_manager.save_presentation(prs_id, sample_file)
        print(f"✅ Presentation saved as: {sample_file}")
        
        # Step 3: Generate screenshots
        print("\n📸 Step 3: Generating screenshots...")
        with tempfile.TemporaryDirectory() as temp_dir:
            screenshot_paths = await ppt_manager.screenshot_slides_async(
                sample_file, temp_dir, "PNG", 1280, 720
            )
            print(f"✅ Generated {len(screenshot_paths)} screenshots:")
            for i, path in enumerate(screenshot_paths):
                if os.path.exists(path):
                    size_kb = os.path.getsize(path) / 1024
                    print(f"   📸 Slide {i+1}: {os.path.basename(path)} ({size_kb:.1f} KB)")
            
            # Step 4: Run critique analysis
            print("\n🔍 Step 4: Running presentation critique...")
            critique_results = await ppt_manager.critique_presentation_async(
                sample_file, "comprehensive", include_screenshots=True, output_dir=temp_dir
            )
            
            # Step 5: Display results
            print("\n📊 Step 5: Critique Results")
            print("-" * 30)
            
            summary = critique_results["summary"]
            print(f"Overall Assessment: {summary['assessment']} ({summary['overall_score']}/100)")
            print(f"Total Slides: {summary['total_slides']}")
            print(f"Critical Issues: {summary['critical_issues']}")
            print(f"Warnings: {summary['warnings']}")
            print(f"Recommendations: {summary['recommendations']}")
            
            # Show analysis breakdown
            if "detailed_analysis" in critique_results:
                print(f"\n📋 Analysis Breakdown:")
                for category, analysis in critique_results["detailed_analysis"].items():
                    score = analysis.get("score", 0)
                    issues = len(analysis.get("issues", []))
                    print(f"   • {category.title()}: {score}/100 ({issues} issues)")
            
            # Show sample issues (if any)
            if critique_results["issues"]:
                print(f"\n🚨 Sample Issues:")
                for issue in critique_results["issues"][:3]:
                    emoji = "🔴" if issue["type"] == "critical" else "⚠️"
                    slide_info = f"Slide {issue['slide']}" if issue['slide'] != 'global' else "Global"
                    print(f"   {emoji} {slide_info}: {issue['issue']}")
            
            # Show strengths
            if critique_results["strengths"]:
                print(f"\n✅ Strengths:")
                for strength in critique_results["strengths"][:3]:
                    print(f"   • {strength}")
            
            # Show top recommendations
            if critique_results["recommendations"]:
                print(f"\n💡 Recommendations:")
                unique_recs = list(set(critique_results["recommendations"]))
                for rec in unique_recs[:3]:
                    print(f"   • {rec}")
            
            print(f"\n📸 Screenshot Integration:")
            if critique_results.get("screenshots"):
                print(f"   ✅ {len(critique_results['screenshots'])} screenshots generated and linked")
                print(f"   📁 Location: {temp_dir}")
            else:
                print(f"   ❌ No screenshots generated")
            
        # Step 6: Demonstrate specific critique types
        print("\n🎯 Step 6: Testing specific critique types...")
        critique_types = ["design", "content", "accessibility", "technical"]
        
        for critique_type in critique_types:
            critique_results = await ppt_manager.critique_presentation_async(
                sample_file, critique_type, include_screenshots=False
            )
            score = critique_results["summary"]["overall_score"]
            issues = len(critique_results["issues"])
            print(f"   • {critique_type.title()}: {score}/100 ({issues} issues)")
        
        print("\n🎉 Example completed successfully!")
        print("\n📝 Next Steps:")
        print("   • Try with your own PowerPoint files")
        print("   • Experiment with different critique types")
        print("   • Use screenshots for AI vision analysis")
        print("   • Integrate into your workflow automation")
        
    except Exception as e:
        print(f"\n❌ Error occurred: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Cleanup
        if 'prs_id' in locals() and prs_id in ppt_manager.presentations:
            del ppt_manager.presentations[prs_id]
        if 'sample_file' in locals() and os.path.exists(sample_file):
            os.remove(sample_file)
        ppt_manager.cleanup()
        print(f"\n🧹 Cleanup completed")

if __name__ == "__main__":
    asyncio.run(main()) 