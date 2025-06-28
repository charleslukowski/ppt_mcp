# PowerPoint MCP Server Improvement Plan

This plan addresses the key issues making your MCP server responses "crummy" and provides actionable steps to dramatically improve usefulness and reliability.

## 🎯 Executive Summary

**Primary Issues Identified:**
- Complex 3-4k token prompt walls causing cognitive overload
- Poor return content types (mostly TextContent)
- No input validation or error handling
- Missing post-processing and quality checks
- No instrumentation or feedback loops

**Expected Impact**: Implementing even 2-3 of these changes should improve response quality by 60-80%.

---

## 🚀 Priority 1: Critical Quick Wins

### 1.1 Streamline Prompt Architecture

**Current Problem**: Your server chains `POWERPOINT_SERVER_SYSTEM_PROMPT` → category prompts → tool prompts → contextual guidance, creating 3-4k token walls.

**Solution**: Implement a lean prompt strategy:

```python
# Replace the current complex system with this approach:
CORE_SYSTEM_PROMPT = """You are a PowerPoint expert focused on professional, accessible presentations. 
Key principles: clean layouts, consistent styling, clear hierarchy, readable fonts."""  # ~400 tokens max

# Dynamic tool-specific injection
def get_focused_prompt(tool_name: str, **kwargs) -> str:
    base = CORE_SYSTEM_PROMPT
    tool_specific = FOCUSED_TOOL_GUIDANCE.get(tool_name, "")
    dynamic_context = f"Success criteria: {get_success_criteria(tool_name, **kwargs)}"
    return f"{base}\n\n{tool_specific}\n\n{dynamic_context}"
```

**Implementation Steps:**
1. Create `FOCUSED_TOOL_GUIDANCE` dict with 2-3 sentence tool-specific tips
2. Add `get_success_criteria()` function that returns "You are done when..." bullets
3. Replace `get_system_prompt()` method with `get_focused_prompt()`

### 1.2 Add Schema Examples

**Current Problem**: No examples in tool definitions, leading to format confusion.

**Solution**: Add examples directly to each tool schema:

```python
# In your tool definitions:
{
    "name": "add_chart",
    "description": "Create a data-driven chart",
    "inputSchema": {
        # ... existing schema ...
        "examples": [
            {
                "chart_type": "column",
                "categories": ["Q1", "Q2", "Q3"],
                "series_data": {"Revenue": [100, 150, 120]}
            }
        ]
    }
}
```

### 1.3 Return Rich Content Types

**Current Problem**: Everything returns `TextContent`, user sees no visual feedback.

**Immediate Fix**: Update key tools to return `EmbeddedResource` or `ImageContent`:

```python
# After successful add_image:
return [
    TextContent(text=f"✅ Image added to slide {slide_index}"),
    EmbeddedResource(
        uri=f"file://{image_path}",
        mimeType="image/png"
    )
]

# After save_presentation:
return [
    TextContent(text=f"✅ Presentation saved: {file_path}"),
    EmbeddedResource(
        uri=f"file://{file_path}",
        mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
]
```

---

## 🛡️ Priority 2: Input Validation & Error Prevention

### 2.1 Pydantic Validation Layer

**Current Problem**: Bad inputs reach the LLM, causing confused responses.

**Solution**: Add input validation before prompt assembly:

```python
from pydantic import BaseModel, validator
from typing import Literal

class AddChartRequest(BaseModel):
    presentation_id: str
    slide_index: int
    chart_type: Literal["column", "bar", "line", "pie"]
    categories: List[str]
    series_data: Dict[str, List[float]]
    
    @validator('categories')
    def categories_not_empty(cls, v):
        if not v:
            raise ValueError("Categories cannot be empty")
        return v

# In your tool handler:
@server.call_tool()
async def handle_call_tool(name: str, arguments: Dict[str, Any]):
    if name == "add_chart":
        try:
            validated_args = AddChartRequest(**arguments)
            result = manager.add_chart(**validated_args.dict())
            return format_success_response(result)
        except ValidationError as e:
            return [TextContent(text=f"❌ Invalid input: {e}")]
```

### 2.2 Async Isolation

**Current Problem**: Heavy operations (save, COM calls) can freeze the event loop.

**Solution**: Wrap heavy operations:

```python
async def save_presentation_async(self, prs_id: str, file_path: str) -> bool:
    return await asyncio.to_thread(self._save_presentation_sync, prs_id, file_path)

def _save_presentation_sync(self, prs_id: str, file_path: str) -> bool:
    # Existing synchronous save logic
    pass
```

---

## 🎨 Priority 3: Post-Processing Quality Checks

### 3.1 PowerPoint-Specific Cleanup

**Known Issues & Fixes:**

| Issue | Detection | Fix |
|-------|-----------|-----|
| Green rectangles covering slides | Check for shapes with fill but no content | Skip fill on placeholder shapes |
| Lost bullet formatting | Text boxes missing paragraph levels | Set `paragraph.level = 0` |
| Blank first slides | Slide with no shapes | Delete empty slides before returning |
| Overlapping elements | Check shape boundaries | Apply minimum spacing rules |

**Implementation:**

```python
def post_process_slide(self, prs_id: str, slide_index: int):
    """Apply quality checks and fixes after operations"""
    prs = self.presentations[prs_id]
    slide = prs.slides[slide_index]
    
    # Fix 1: Clean up placeholder fills
    for shape in slide.shapes:
        if shape.is_placeholder and hasattr(shape, 'fill'):
            shape.fill.background()
    
    # Fix 2: Ensure proper bulleting
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame'):
            for paragraph in shape.text_frame.paragraphs:
                if not paragraph.level:
                    paragraph.level = 0
    
    # Fix 3: Remove empty slides
    if len(slide.shapes) == 0:
        self.delete_slide(prs_id, slide_index)
```

---

## 📊 Priority 4: User Experience Enhancements

### 4.1 Result Quality Knobs

**Solution**: Add optional preferences to every tool:

```python
class ResultPreferences(BaseModel):
    verbosity: Literal["brief", "normal", "verbose"] = "normal"
    tone: Literal["executive", "technical", "friendly"] = "professional"
    include_preview: bool = True

# In tool responses:
def format_response(result: Any, preferences: ResultPreferences):
    if preferences.verbosity == "brief":
        return [TextContent(text="✅ Done")]
    elif preferences.verbosity == "verbose":
        return [TextContent(text=detailed_explanation)]
    # ... etc
```

### 4.2 Better Success Indicators

**Current**: Generic "operation completed" messages  
**Improved**: Specific, actionable feedback:

```python
# Instead of: "Chart added successfully"
# Return: "✅ Column chart added to slide 2 showing Q1-Q4 data with 2 series (Revenue, Profit)"

def format_chart_success(chart_type: str, slide_index: int, categories: List[str], series_names: List[str]) -> str:
    return f"✅ {chart_type.title()} chart added to slide {slide_index+1} showing {len(categories)} periods with {len(series_names)} series ({', '.join(series_names)})"
```

---

## 📈 Priority 5: Instrumentation & Continuous Improvement

### 5.1 Operation Logging

**Implementation:**

```python
import time
import json
from datetime import datetime

class OperationLogger:
    def __init__(self):
        self.log_file = "mcp_operations.jsonl"
    
    def log_operation(self, tool: str, args: dict, tokens_in: int, tokens_out: int, 
                     latency_ms: int, success: bool, error: str = None):
        log_entry = {
            "timestamp": datetime.utcnow().isoformat(),
            "tool": tool,
            "args_hash": hash(str(sorted(args.items()))),
            "tokens_in": tokens_in,
            "tokens_out": tokens_out,
            "latency_ms": latency_ms,
            "success": success,
            "error": error
        }
        
        with open(self.log_file, "a") as f:
            f.write(json.dumps(log_entry) + "\n")

# Usage in tool handlers:
start_time = time.time()
try:
    result = manager.add_chart(**args)
    logger.log_operation("add_chart", args, tokens_in, tokens_out, 
                        int((time.time() - start_time) * 1000), True)
except Exception as e:
    logger.log_operation("add_chart", args, tokens_in, tokens_out,
                        int((time.time() - start_time) * 1000), False, str(e))
```

### 5.2 Quality Regression Tests

**Create test cases for prompt changes:**

```python
def test_add_chart_prompt_quality():
    """Ensure prompt changes don't break chart generation"""
    prompt = get_focused_prompt("add_chart", chart_type="column", categories=["A", "B"])
    
    # Quality heuristics
    assert len(prompt) < 1000  # Not too verbose
    assert "column" in prompt.lower()  # Includes specifics
    assert "success criteria" in prompt.lower()  # Has success criteria
    assert not any(word in prompt.lower() for word in ["error", "failed", "invalid"])  # Positive tone
```

---

## 📋 Implementation Checklist

### Week 1: Foundation (Biggest Impact)
- [ ] Implement lean prompt architecture (1.1)
- [ ] Add schema examples to top 5 tools (1.2)
- [ ] Update `add_image`, `save_presentation`, `add_chart` to return rich content (1.3)
- [ ] Add success criteria footers to prompts

### Week 2: Validation & Error Handling
- [ ] Add Pydantic models for top 5 tools (2.1)
- [ ] Implement async isolation for heavy operations (2.2)
- [ ] Add post-processing cleanup for known issues (3.1)

### Week 3: Polish & Monitoring
- [ ] Add result preferences system (4.1)
- [ ] Improve success messages (4.2)
- [ ] Implement operation logging (5.1)
- [ ] Create quality regression tests (5.2)

### Week 4: Fine-tuning
- [ ] Analyze logs for common failure patterns
- [ ] Optimize prompts based on real usage data
- [ ] Add user feedback collection mechanism

---

## 🎯 Success Metrics

**Before/After Comparison:**
- **Prompt Length**: 3000+ tokens → <800 tokens
- **User Satisfaction**: Track via feedback collection
- **Error Rate**: Measure validation catches vs LLM confusion
- **Response Time**: Monitor latency improvements
- **Feature Adoption**: Track which tools are used most

**Quality Indicators:**
- Fewer "I don't understand" responses
- More successful first-try operations  
- Reduced back-and-forth clarifications
- Positive user feedback on visual outputs

---

## 💡 Advanced Improvements (Phase 2)

Once core issues are resolved, consider:

1. **Intelligent Defaults**: Learn user preferences over time
2. **Template Library**: Pre-built professional templates
3. **Content Suggestions**: AI-powered layout and design recommendations
4. **Batch Operations**: Multi-slide operations in single calls
5. **Integration Hooks**: Connect with external data sources
6. **Preview Generation**: Thumbnail previews of changes
7. **Undo/Redo**: Operation history and rollback capabilities

---

**Next Steps**: Start with Week 1 tasks. Even implementing the lean prompt architecture alone should show immediate improvement. Monitor logs and user feedback to prioritize subsequent changes.

*Ready to implement? Let me know which specific area you'd like code examples for!*