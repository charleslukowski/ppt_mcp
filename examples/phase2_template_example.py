#!/usr/bin/env python3
"""
Phase 2: Content Automation & Templates - Comprehensive Example

This example demonstrates the complete Phase 2 template system including:
- Template creation with placeholders
- Variable substitution with nested data
- Conditional logic for dynamic slides
- Bulk presentation generation
- Content updates and data mapping
"""

import json
import tempfile
import os
from pathlib import Path

# Example: Creating and using templates with the PowerPoint MCP Server

def create_monthly_report_template():
    """Example template configuration for monthly business reports"""
    
    template_config = {
        "name": "Monthly Business Report Template",
        "description": "Comprehensive template for monthly business reporting with dynamic content and conditional sections",
        "slides": [
            {
                "layout_type": "title_slide",
                "elements": [
                    {
                        "type": "text",
                        "content": "{{company.name}} Monthly Report",
                        "position": {"left": 1, "top": 1.5, "width": 8, "height": 1.5},
                        "formatting": {"font_size": 36, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "{{report.period}} {{report.year}}",
                        "position": {"left": 1, "top": 3, "width": 8, "height": 1},
                        "formatting": {"font_size": 24, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Prepared by: {{report.author}}",
                        "position": {"left": 1, "top": 5, "width": 8, "height": 0.8},
                        "formatting": {"font_size": 16, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Generated on: {{report.date}}",
                        "position": {"left": 1, "top": 5.8, "width": 8, "height": 0.8},
                        "formatting": {"font_size": 14, "bold": False}
                    }
                ]
            },
            {
                "layout_type": "executive_summary",
                "elements": [
                    {
                        "type": "text",
                        "content": "Executive Summary",
                        "position": {"left": 1, "top": 1, "width": 8, "height": 1},
                        "formatting": {"font_size": 28, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "{{summary.overview}}",
                        "position": {"left": 1, "top": 2, "width": 8, "height": 2},
                        "formatting": {"font_size": 16, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Key Highlights:",
                        "position": {"left": 1, "top": 4.5, "width": 8, "height": 0.5},
                        "formatting": {"font_size": 18, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "• {{summary.highlight1}}\n• {{summary.highlight2}}\n• {{summary.highlight3}}",
                        "position": {"left": 1, "top": 5, "width": 8, "height": 1.5},
                        "formatting": {"font_size": 14, "bold": False}
                    }
                ]
            },
            {
                "layout_type": "performance_metrics",
                "elements": [
                    {
                        "type": "text",
                        "content": "Performance Metrics",
                        "position": {"left": 1, "top": 1, "width": 8, "height": 1},
                        "formatting": {"font_size": 28, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "Revenue: ${{metrics.revenue}}M ({{metrics.revenue_change}}% vs last month)",
                        "position": {"left": 1, "top": 2.5, "width": 8, "height": 0.5},
                        "formatting": {"font_size": 16, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Customers: {{metrics.customers}} ({{metrics.customer_change}}% vs last month)",
                        "position": {"left": 1, "top": 3, "width": 8, "height": 0.5},
                        "formatting": {"font_size": 16, "bold": False}
                    },
                    {
                        "type": "chart",
                        "chart_type": "column",
                        "data": {
                            "categories": "metrics.chart.categories",
                            "series": "metrics.chart.series"
                        },
                        "position": {"left": 1, "top": 4, "width": 8, "height": 3}
                    }
                ],
                "conditional_logic": {
                    "if": {
                        "field": "metrics.include_charts",
                        "operator": "equals",
                        "value": True
                    }
                }
            },
            {
                "layout_type": "success_slide",
                "elements": [
                    {
                        "type": "text",
                        "content": "🎉 Outstanding Performance!",
                        "position": {"left": 1, "top": 2, "width": 8, "height": 1},
                        "formatting": {"font_size": 32, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "Revenue exceeded target by {{metrics.revenue_over_target}}%",
                        "position": {"left": 1, "top": 3.5, "width": 8, "height": 1},
                        "formatting": {"font_size": 20, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "{{success.message}}",
                        "position": {"left": 1, "top": 5, "width": 8, "height": 2},
                        "formatting": {"font_size": 16, "bold": False}
                    }
                ],
                "conditional_logic": {
                    "if": {
                        "field": "metrics.revenue",
                        "operator": "greater_than",
                        "value": 100
                    }
                }
            },
            {
                "layout_type": "improvement_needed",
                "elements": [
                    {
                        "type": "text",
                        "content": "📈 Areas for Improvement",
                        "position": {"left": 1, "top": 2, "width": 8, "height": 1},
                        "formatting": {"font_size": 28, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "Revenue: ${{metrics.revenue}}M (Below target of $100M)",
                        "position": {"left": 1, "top": 3.5, "width": 8, "height": 0.5},
                        "formatting": {"font_size": 18, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Action Plan:\n{{improvement.action_plan}}",
                        "position": {"left": 1, "top": 4.5, "width": 8, "height": 2},
                        "formatting": {"font_size": 16, "bold": False}
                    }
                ],
                "conditional_logic": {
                    "if": {
                        "field": "metrics.revenue",
                        "operator": "less_than",
                        "value": 100
                    }
                }
            },
            {
                "layout_type": "next_steps",
                "elements": [
                    {
                        "type": "text",
                        "content": "Next Steps & Action Items",
                        "position": {"left": 1, "top": 1, "width": 8, "height": 1},
                        "formatting": {"font_size": 28, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "Priority Actions for {{next_month.name}}:",
                        "position": {"left": 1, "top": 2, "width": 8, "height": 0.5},
                        "formatting": {"font_size": 18, "bold": True}
                    },
                    {
                        "type": "text",
                        "content": "1. {{next_steps.action1}}\n2. {{next_steps.action2}}\n3. {{next_steps.action3}}",
                        "position": {"left": 1, "top": 2.8, "width": 8, "height": 2},
                        "formatting": {"font_size": 16, "bold": False}
                    },
                    {
                        "type": "text",
                        "content": "Target Goals:\n• Revenue: ${{next_month.revenue_target}}M\n• New Customers: {{next_month.customer_target}}",
                        "position": {"left": 1, "top": 5, "width": 8, "height": 1.5},
                        "formatting": {"font_size": 14, "bold": False}
                    }
                ]
            }
        ]
    }
    
    return template_config

def create_sample_data_high_performance():
    """Sample data for a high-performing month (will show success slide)"""
    
    return {
        "company": {
            "name": "TechCorp Industries",
            "department": "Sales & Marketing"
        },
        "report": {
            "period": "November",
            "year": "2024",
            "author": "Sarah Johnson",
            "date": "December 1, 2024"
        },
        "summary": {
            "overview": "November was an exceptional month with record-breaking performance across all key metrics. The team exceeded targets through strategic client acquisition and successful product launches.",
            "highlight1": "Revenue grew 25% month-over-month to $125M",
            "highlight2": "Acquired 150 new enterprise customers",
            "highlight3": "Product satisfaction scores reached 98%"
        },
        "metrics": {
            "revenue": 125,
            "revenue_change": 25,
            "revenue_over_target": 25,
            "customers": 1250,
            "customer_change": 15,
            "include_charts": True,
            "chart": {
                "categories": ["Sales", "Marketing", "Customer Success", "Product"],
                "series": {
                    "November": [125, 95, 88, 92],
                    "Target": [100, 85, 85, 85],
                    "October": [100, 88, 85, 88]
                }
            }
        },
        "success": {
            "message": "Congratulations to the entire team for this outstanding achievement! This performance positions us perfectly for Q4 success and sets a strong foundation for next year's growth targets."
        },
        "next_steps": {
            "action1": "Expand successful marketing campaigns to new regions",
            "action2": "Onboard 50% of new customers within 30 days",
            "action3": "Launch advanced product features based on customer feedback"
        },
        "next_month": {
            "name": "December",
            "revenue_target": 130,
            "customer_target": 200
        }
    }

def create_sample_data_low_performance():
    """Sample data for an underperforming month (will show improvement slide)"""
    
    return {
        "company": {
            "name": "TechCorp Industries",
            "department": "Sales & Marketing"
        },
        "report": {
            "period": "September", 
            "year": "2024",
            "author": "Michael Chen",
            "date": "October 1, 2024"
        },
        "summary": {
            "overview": "September presented challenges with market headwinds and increased competition. While we faced setbacks in revenue, we made important improvements in customer satisfaction and operational efficiency.",
            "highlight1": "Improved customer satisfaction scores by 15%",
            "highlight2": "Reduced operational costs by 8%",
            "highlight3": "Launched new customer support initiatives"
        },
        "metrics": {
            "revenue": 75,
            "revenue_change": -12,
            "customers": 980,
            "customer_change": -5,
            "include_charts": True,
            "chart": {
                "categories": ["Sales", "Marketing", "Customer Success", "Product"],
                "series": {
                    "September": [75, 70, 92, 85],
                    "Target": [100, 85, 85, 85],
                    "August": [85, 78, 88, 82]
                }
            }
        },
        "improvement": {
            "action_plan": "• Accelerate lead generation efforts with new marketing campaigns\n• Focus on high-value client retention programs\n• Implement pricing optimization strategies\n• Enhance sales team training and support"
        },
        "next_steps": {
            "action1": "Launch targeted marketing campaign for Q4 push",
            "action2": "Implement customer win-back program",
            "action3": "Optimize pricing structure for competitive advantage"
        },
        "next_month": {
            "name": "October",
            "revenue_target": 95,
            "customer_target": 100
        }
    }

def create_bulk_generation_datasets():
    """Create multiple datasets for bulk presentation generation"""
    
    datasets = []
    
    # Q1 Data
    datasets.append({
        "company": {"name": "TechCorp Industries"},
        "report": {"period": "Q1", "year": "2024", "author": "Alice Brown", "date": "April 1, 2024"},
        "summary": {
            "overview": "Q1 established strong momentum with steady growth and successful market expansion.",
            "highlight1": "Entered 3 new international markets",
            "highlight2": "Launched 2 major product features", 
            "highlight3": "Achieved 95% customer retention rate"
        },
        "metrics": {"revenue": 110, "revenue_change": 10, "customers": 1100, "customer_change": 8, "include_charts": True,
                   "chart": {"categories": ["Sales", "Marketing", "Support"], "series": {"Q1": [110, 88, 95], "Target": [100, 85, 90]}}},
        "success": {"message": "Excellent start to the year with solid fundamentals in place."},
        "next_steps": {"action1": "Scale international operations", "action2": "Expand product portfolio", "action3": "Strengthen market presence"},
        "next_month": {"name": "Q2", "revenue_target": 115, "customer_target": 150}
    })
    
    # Q2 Data  
    datasets.append({
        "company": {"name": "TechCorp Industries"},
        "report": {"period": "Q2", "year": "2024", "author": "Bob Martinez", "date": "July 1, 2024"},
        "summary": {
            "overview": "Q2 accelerated growth with breakthrough achievements in customer acquisition.",
            "highlight1": "Record customer acquisition of 300 new clients",
            "highlight2": "Revenue growth of 20% quarter-over-quarter",
            "highlight3": "Product innovation leading to 25% efficiency gains"
        },
        "metrics": {"revenue": 135, "revenue_change": 20, "customers": 1400, "customer_change": 25, "include_charts": True,
                   "chart": {"categories": ["Sales", "Marketing", "Support"], "series": {"Q2": [135, 95, 88], "Target": [115, 88, 90]}}},
        "success": {"message": "Outstanding Q2 performance exceeding all expectations and setting new company records."},
        "next_steps": {"action1": "Optimize customer onboarding", "action2": "Expand customer success team", "action3": "Prepare for Q3 scaling"},
        "next_month": {"name": "Q3", "revenue_target": 140, "customer_target": 200}
    })
    
    # Q3 Data
    datasets.append({
        "company": {"name": "TechCorp Industries"},
        "report": {"period": "Q3", "year": "2024", "author": "Carol Zhang", "date": "October 1, 2024"},
        "summary": {
            "overview": "Q3 maintained strong momentum despite market challenges, with focus on sustainable growth.",
            "highlight1": "Maintained 95% customer satisfaction rate",
            "highlight2": "Achieved cost optimization targets",
            "highlight3": "Expanded into emerging market segments"
        },
        "metrics": {"revenue": 88, "revenue_change": -5, "customers": 1350, "customer_change": -3, "include_charts": True,
                   "chart": {"categories": ["Sales", "Marketing", "Support"], "series": {"Q3": [88, 85, 92], "Target": [140, 90, 88]}}},
        "improvement": {"action_plan": "• Focus on premium customer segments\n• Accelerate innovation pipeline\n• Strengthen competitive positioning\n• Optimize market penetration strategies"},
        "next_steps": {"action1": "Launch premium service tier", "action2": "Accelerate product development", "action3": "Strengthen competitive analysis"},
        "next_month": {"name": "Q4", "revenue_target": 120, "customer_target": 180}
    })
    
    return datasets

def demonstrate_data_source_mapping():
    """Example of data source configuration for template integration"""
    
    # Example JSON data source
    sample_data_source = {
        "company_info": {
            "name": "TechCorp Industries",
            "industry": "Technology",
            "founded": 2015,
            "headquarters": "San Francisco, CA"
        },
        "monthly_reports": [
            {
                "month": "November",
                "year": 2024,
                "metrics": {
                    "revenue": 125,
                    "customers": 1250,
                    "satisfaction": 4.8
                },
                "summary": "Exceptional performance with record results"
            },
            {
                "month": "October", 
                "year": 2024,
                "metrics": {
                    "revenue": 100,
                    "customers": 1100,
                    "satisfaction": 4.6
                },
                "summary": "Steady growth with solid fundamentals"
            }
        ]
    }
    
    # Data source configuration
    source_config = {
        "type": "json",
        "source": "monthly_reports.json",
        "mapping": {
            "company.name": "company_info.name",
            "report.period": "monthly_reports.0.month",
            "report.year": "monthly_reports.0.year",
            "metrics.revenue": "monthly_reports.0.metrics.revenue",
            "metrics.customers": "monthly_reports.0.metrics.customers",
            "summary.overview": "monthly_reports.0.summary"
        },
        "refresh_interval": 3600
    }
    
    return sample_data_source, source_config

def main():
    """Main function demonstrating Phase 2 capabilities"""
    
    print("🤖 Phase 2: Content Automation & Templates - Comprehensive Example")
    print("=" * 70)
    
    # 1. Template Configuration
    print("\n1. Template Configuration")
    print("-" * 30)
    
    template_config = create_monthly_report_template()
    print(f"Template: {template_config['name']}")
    print(f"Slides: {len(template_config['slides'])}")
    print(f"Description: {template_config['description']}")
    
    # 2. Sample Data Examples
    print("\n2. Sample Data Examples")
    print("-" * 30)
    
    high_perf_data = create_sample_data_high_performance()
    low_perf_data = create_sample_data_low_performance()
    
    print(f"High Performance Data: Revenue ${high_perf_data['metrics']['revenue']}M")
    print(f"Low Performance Data: Revenue ${low_perf_data['metrics']['revenue']}M")
    print("Conditional logic will show different slides based on performance")
    
    # 3. Bulk Generation Datasets
    print("\n3. Bulk Generation Datasets")
    print("-" * 30)
    
    bulk_datasets = create_bulk_generation_datasets()
    print(f"Quarterly Reports: {len(bulk_datasets)} datasets")
    for i, dataset in enumerate(bulk_datasets):
        revenue = dataset['metrics']['revenue']
        period = dataset['report']['period']
        print(f"  {period}: ${revenue}M revenue")
    
    # 4. Data Source Integration
    print("\n4. Data Source Integration")
    print("-" * 30)
    
    data_source, source_config = demonstrate_data_source_mapping()
    print(f"Data Source Type: {source_config['type']}")
    print(f"Mappings: {len(source_config['mapping'])} field mappings")
    print(f"Refresh Interval: {source_config['refresh_interval']} seconds")
    
    # 5. Usage Examples
    print("\n5. Usage Examples")
    print("-" * 30)
    
    print("Example MCP tool calls:")
    print()
    
    # Create template
    print("# Create template")
    print('template_id = await client.call_tool("create_template", {')
    print('    "template_config": template_config')
    print('})')
    print()
    
    # Apply template
    print("# Apply template with high performance data")
    print('prs_id = await client.call_tool("apply_template", {')
    print('    "template_id": template_id,')
    print('    "data": high_performance_data')
    print('})')
    print()
    
    # Bulk generation
    print("# Bulk generate quarterly reports")
    print('presentations = await client.call_tool("bulk_generate_presentations", {')
    print('    "template_id": template_id,')
    print('    "data_sets": quarterly_datasets,')
    print('    "output_config": {"auto_save": True, "output_path": "/reports"}')
    print('})')
    print()
    
    # Update content
    print("# Update existing presentation")
    print('await client.call_tool("update_template_content", {')
    print('    "presentation_id": prs_id,')
    print('    "updates": {')
    print('        "0": {"report.author": "Updated Author"},')
    print('        "1": {"summary.overview": "Updated summary"}')
    print('    }')
    print('})')
    
    # 6. Template Features Summary
    print("\n6. Template Features Summary")
    print("-" * 30)
    
    print("✅ Variable Substitution: {{company.name}}, {{metrics.revenue}}")
    print("✅ Nested Data Access: {{company.department.manager}}")
    print("✅ Conditional Logic: Show/hide slides based on performance")
    print("✅ Multiple Element Types: Text, images, charts")
    print("✅ Flexible Positioning: Left, top, width, height control")
    print("✅ Text Formatting: Font size, bold, italic options")
    print("✅ Chart Integration: Dynamic chart data from variables")
    print("✅ Bulk Generation: Multiple presentations from datasets")
    print("✅ Content Updates: Modify existing presentations")
    print("✅ Data Source Mapping: JSON, CSV, Excel, API support")
    
    print("\n🎉 Phase 2 implementation enables powerful content automation!")
    print("Ready for integration with existing Phase 1 & 4 features.")

if __name__ == "__main__":
    main() 