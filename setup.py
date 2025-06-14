#!/usr/bin/env python3
"""
Setup script for PowerPoint MCP Server

This script allows for proper installation and distribution of the PowerPoint MCP server.
"""

from setuptools import setup, find_packages
import os

# Read the README file for long description
def read_readme():
    readme_path = os.path.join(os.path.dirname(__file__), 'README.md')
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            return f.read()
    return "PowerPoint MCP Server - A comprehensive MCP server for PowerPoint manipulation"

# Read requirements from requirements.txt
def read_requirements():
    requirements_path = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    if os.path.exists(requirements_path):
        with open(requirements_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip() and not line.startswith('#')]
    return [
        'python-pptx>=0.6.21',
        'mcp>=1.0.0',
        'asyncio-mqtt>=0.11.0',
        'typing-extensions>=4.0.0'
    ]

setup(
    name="powerpoint-mcp-server",
    version="1.0.0",
    author="PowerPoint MCP Server Team",
    author_email="contact@example.com",
    description="A comprehensive Model Context Protocol server for PowerPoint presentation manipulation",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/your-username/powerpoint-mcp-server",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Communications",
        "Topic :: Internet :: WWW/HTTP :: Dynamic Content",
    ],
    python_requires=">=3.8",
    install_requires=read_requirements(),
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-asyncio>=0.21.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
        ],
        "test": [
            "pytest>=7.0.0",
            "pytest-asyncio>=0.21.0",
            "pytest-cov>=4.0.0",
        ]
    },
    entry_points={
        "console_scripts": [
            "powerpoint-mcp-server=powerpoint_mcp_server:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["*.md", "*.txt", "*.json"],
    },
    keywords=[
        "powerpoint",
        "mcp",
        "model-context-protocol",
        "presentation",
        "automation",
        "python-pptx",
        "cursor",
        "ai-assistant"
    ],
    project_urls={
        "Bug Reports": "https://github.com/your-username/powerpoint-mcp-server/issues",
        "Source": "https://github.com/your-username/powerpoint-mcp-server",
        "Documentation": "https://github.com/your-username/powerpoint-mcp-server#readme",
    },
) 