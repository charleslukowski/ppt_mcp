#!/usr/bin/env python3
"""
Minimal MCP server test to identify the exact issue.
"""

import json
import subprocess
import sys
import time

def test_basic_mcp():
    """Test basic MCP protocol communication"""
    
    print("🔧 Basic MCP Protocol Test...")
    
    # Start server
    process = subprocess.Popen(
        [sys.executable, "powerpoint_mcp_server.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=0  # Unbuffered
    )
    
    try:
        # Test 1: Initialize (this worked in our previous test)
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }
        
        print("📤 Sending initialize...")
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        
        # Read init response
        init_response = process.stdout.readline()
        print(f"📥 Init response: {init_response.strip()}")
        
        # Test 2: Try tools/list with empty params
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }
        
        print("\n📤 Sending tools/list (empty params)...")
        process.stdin.write(json.dumps(tools_request) + "\n")
        process.stdin.flush()
        
        # Read tools response
        tools_response = process.stdout.readline()
        print(f"📥 Tools response: {tools_response.strip()}")
        
        # Test 3: Try tools/list with no params at all
        tools_request_no_params = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/list"
        }
        
        print("\n📤 Sending tools/list (no params)...")
        process.stdin.write(json.dumps(tools_request_no_params) + "\n")
        process.stdin.flush()
        
        # Read tools response
        tools_response_2 = process.stdout.readline()
        print(f"📥 Tools response 2: {tools_response_2.strip()}")
        
        # Check stderr for any errors
        time.sleep(0.1)  # Give a moment for any stderr output
        if process.stderr:
            stderr_content = process.stderr.read()
            if stderr_content:
                print(f"\n🚨 Stderr output: {stderr_content}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
    
    finally:
        process.terminate()
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()

if __name__ == "__main__":
    test_basic_mcp() 