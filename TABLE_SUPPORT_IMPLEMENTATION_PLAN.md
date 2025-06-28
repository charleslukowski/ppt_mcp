# PowerPoint MCP Server - Table Support Implementation Plan

## Overview

This document outlines the comprehensive plan for adding table support to the PowerPoint MCP Server. The implementation will add full table creation, editing, deletion, and styling capabilities while maintaining compatibility with the existing stable architecture.

## Current State Analysis

### Existing Capabilities
- ✅ Text boxes with comprehensive formatting
- ✅ Images and charts
- ✅ Basic shape manipulation (delete, clear)
- ✅ File operations (save/load)
- ✅ Content extraction and information retrieval
- ✅ Slide background management

### Missing Capabilities
- ❌ Table creation and management
- ❌ Cell content editing
- ❌ Table styling and formatting
- ❌ Table structure modification
- ❌ Table-specific information extraction

## Implementation Strategy

### Core Philosophy
- **Incremental Development**: Implement in phases to maintain stability
- **Consistency**: Follow existing patterns for validation, error handling, and success messages
- **Comprehensive Coverage**: Support all major table operations
- **User-Friendly**: Provide clear feedback and intuitive parameter naming

## New Tools Design

### 1. `add_table` - Table Creation
**Purpose**: Create a new table on a slide with specified dimensions

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required, ≥0)",
  "rows": "integer (required, 1-50)",
  "cols": "integer (required, 1-20)", 
  "left": "number (optional, default: 1)",
  "top": "number (optional, default: 1)",
  "width": "number (optional, default: 8)",
  "height": "number (optional, default: 4)",
  "header_row": "boolean (optional, default: false)"
}
```

**Success Message**: `📊 Added {rows}×{cols} table to slide {slide_index + 1}`

### 2. `set_table_cell` - Cell Content Management
**Purpose**: Set text content and basic formatting for individual cells

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required)",
  "table_index": "integer (required, ≥0)",
  "row": "integer (required, ≥0)",
  "col": "integer (required, ≥0)",
  "text": "string (required)",
  "font_size": "integer (optional, 8-72)",
  "font_name": "string (optional)",
  "font_color": "string (optional)",
  "bold": "boolean (optional)",
  "italic": "boolean (optional)",
  "underline": "boolean (optional)",
  "text_alignment": "string (optional, left|center|right|justify)"
}
```

**Success Message**: `✅ Updated table {table_index} cell [{row},{col}] on slide {slide_index + 1}: "{text_preview}"`

### 3. `style_table_cell` - Cell Styling
**Purpose**: Apply background colors, borders, and margins to individual cells

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required)",
  "table_index": "integer (required)",
  "row": "integer (required)",
  "col": "integer (required)",
  "fill_color": "string (optional)",
  "border_color": "string (optional)",
  "border_width": "number (optional)",
  "margin_left": "number (optional)",
  "margin_right": "number (optional)",
  "margin_top": "number (optional)",
  "margin_bottom": "number (optional)"
}
```

**Success Message**: `🎨 Styled table {table_index} cell [{row},{col}] on slide {slide_index + 1}`

### 4. `style_table_range` - Bulk Cell Styling
**Purpose**: Apply styling to multiple cells simultaneously

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required)",
  "table_index": "integer (required)",
  "start_row": "integer (required)",
  "start_col": "integer (required)",
  "end_row": "integer (required)",
  "end_col": "integer (required)",
  "fill_color": "string (optional)",
  "border_color": "string (optional)",
  "border_width": "number (optional)",
  "margin_left": "number (optional)",
  "margin_right": "number (optional)",
  "margin_top": "number (optional)",
  "margin_bottom": "number (optional)"
}
```

**Success Message**: `🎨 Styled table {table_index} range [{start_row},{start_col}] to [{end_row},{end_col}] on slide {slide_index + 1}`

### 5. `modify_table_structure` - Structure Modification
**Purpose**: Add or remove rows and columns dynamically

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required)",
  "table_index": "integer (required)",
  "action": "string (required, add_row|delete_row|add_column|delete_column)",
  "index": "integer (required, ≥0)"
}
```

**Success Message**: `🔧 {action} at index {index} for table {table_index} on slide {slide_index + 1}`

### 6. `get_table_info` - Table Inspection
**Purpose**: Retrieve comprehensive information about a table

**Parameters**:
```json
{
  "presentation_id": "string (required)",
  "slide_index": "integer (required)",
  "table_index": "integer (required)"
}
```

**Success Message**: `ℹ️ Table {table_index} info: {rows}×{cols} table with {total_cells} cells`

## Implementation Phases

### Phase 1: Foundation (Week 1)
**Goal**: Basic table creation and infrastructure

**Tasks**:
- [ ] Add table-specific validation functions
- [ ] Implement `add_table` method in PowerPointManager
- [ ] Add `add_table` tool registration
- [ ] Implement tool handler for `add_table`
- [ ] Add table-specific success message formatting
- [ ] Create helper method `_get_table()` for table retrieval
- [ ] Basic testing and validation

**Deliverables**:
- Working `add_table` tool
- Basic table creation with positioning and header styling
- Integration with existing validation framework

### Phase 2: Content Management (Week 2)  
**Goal**: Cell content editing and basic formatting

**Tasks**:
- [ ] Implement `set_table_cell` method
- [ ] Add text formatting support within cells
- [ ] Implement `get_table_info` method
- [ ] Add tool registrations and handlers
- [ ] Enhance `extract_text` to include table content
- [ ] Update `list_slide_content` to identify tables

**Deliverables**:
- Cell text content setting and formatting
- Table information retrieval
- Enhanced content extraction

### Phase 3: Advanced Styling (Week 3)
**Goal**: Comprehensive cell and range styling

**Tasks**:
- [ ] Implement `style_table_cell` method
- [ ] Implement `style_table_range` method
- [ ] Add support for cell backgrounds and borders
- [ ] Implement margin controls
- [ ] Add color validation for table-specific contexts
- [ ] Range operation optimization

**Deliverables**:
- Individual cell styling
- Range-based styling operations
- Comprehensive visual formatting options

### Phase 4: Structure Modification (Week 4)
**Goal**: Dynamic table structure changes

**Tasks**:
- [ ] Research python-pptx limitations for row/column operations
- [ ] Implement `modify_table_structure` method
- [ ] Develop workarounds for python-pptx limitations
- [ ] Add comprehensive error handling for edge cases
- [ ] Performance optimization for large tables
- [ ] Complete integration testing

**Deliverables**:
- Row and column addition/deletion
- Robust error handling
- Performance-optimized operations

## Technical Implementation Details

### Code Structure Extensions

#### 1. Validation Extensions
```python
def validate_table_args(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
    """Extended validation for table-specific operations"""
    
    if tool_name == "add_table":
        rows = arguments.get("rows")
        cols = arguments.get("cols")
        if not isinstance(rows, int) or rows < 1 or rows > 50:
            raise ValueError("rows must be between 1 and 50")
        if not isinstance(cols, int) or cols < 1 or cols > 20:
            raise ValueError("cols must be between 1 and 20")
            
    elif tool_name in ["set_table_cell", "style_table_cell"]:
        table_index = arguments.get("table_index")
        if table_index is None or not isinstance(table_index, int) or table_index < 0:
            raise ValueError("table_index must be a non-negative integer")
        
        row = arguments.get("row")
        col = arguments.get("col") 
        if row is None or not isinstance(row, int) or row < 0:
            raise ValueError("row must be a non-negative integer")
        if col is None or not isinstance(col, int) or col < 0:
            raise ValueError("col must be a non-negative integer")
            
    elif tool_name == "style_table_range":
        # Range validation
        start_row = arguments.get("start_row", 0)
        end_row = arguments.get("end_row", 0)
        start_col = arguments.get("start_col", 0)
        end_col = arguments.get("end_col", 0)
        
        if start_row > end_row:
            raise ValueError("start_row must be <= end_row")
        if start_col > end_col:
            raise ValueError("start_col must be <= end_col")
            
    elif tool_name == "modify_table_structure":
        action = arguments.get("action")
        valid_actions = ["add_row", "delete_row", "add_column", "delete_column"]
        if action not in valid_actions:
            raise ValueError(f"action must be one of: {valid_actions}")
            
        index = arguments.get("index")
        if index is None or not isinstance(index, int) or index < 0:
            raise ValueError("index must be a non-negative integer")
    
    return arguments
```

#### 2. Core PowerPointManager Methods

**Helper Methods**:
```python
def _get_table_shape(self, prs_id: str, slide_index: int, table_index: int):
    """Get table shape object with validation"""
    if prs_id not in self.presentations:
        raise ValueError(f"Presentation {prs_id} not found")
    
    prs = self.presentations[prs_id]
    if slide_index >= len(prs.slides):
        raise ValueError(f"Slide {slide_index} does not exist")
    
    slide = prs.slides[slide_index]
    tables = [shape for shape in slide.shapes if hasattr(shape, 'table')]
    
    if table_index >= len(tables):
        raise ValueError(f"Table {table_index} does not exist (found {len(tables)} tables)")
    
    return tables[table_index]

def _get_table(self, prs_id: str, slide_index: int, table_index: int):
    """Get table object with validation"""
    return self._get_table_shape(prs_id, slide_index, table_index).table
```

#### 3. Enhanced Content Extraction
Update existing methods to handle tables:

```python
# In extract_text method, add table handling:
elif hasattr(shape, 'table'):
    table_text = self._extract_table_text(shape.table, shape_idx)
    if table_text:
        slide_text["text_content"].append(table_text)

def _extract_table_text(self, table, shape_idx):
    """Extract text content from table cells"""
    table_content = []
    for row_idx, row in enumerate(table.rows):
        row_content = []
        for col_idx, cell in enumerate(row.cells):
            if cell.text.strip():
                row_content.append(cell.text.strip())
        if row_content:
            table_content.append(" | ".join(row_content))
    
    if table_content:
        return {
            "shape_index": shape_idx,
            "shape_type": "table",
            "text": "\n".join(table_content),
            "rows": len(table.rows),
            "columns": len(table.columns)
        }
    return None
```

## Usage Examples

### Basic Table Creation
```python
# Create a 3x4 table with header
add_table(
    presentation_id="ppt_0",
    slide_index=1,
    rows=4,
    cols=3,
    header_row=True
)
```

### Cell Content Management
```python
# Set header cells
set_table_cell(
    presentation_id="ppt_0",
    slide_index=1,
    table_index=0,
    row=0,
    col=0,
    text="Product",
    bold=True,
    font_size=14
)

# Set data cells
set_table_cell(
    presentation_id="ppt_0", 
    slide_index=1,
    table_index=0,
    row=1,
    col=0,
    text="Widget A",
    font_size=12
)
```

### Styling Operations
```python
# Style header row
style_table_range(
    presentation_id="ppt_0",
    slide_index=1,
    table_index=0,
    start_row=0,
    start_col=0,
    end_row=0,
    end_col=2,
    fill_color="#4472C4",
    border_color="black",
    border_width=1
)

# Style data area with alternating colors
style_table_range(
    presentation_id="ppt_0",
    slide_index=1,
    table_index=0,
    start_row=1,
    start_col=0,
    end_row=2,
    end_col=2,
    fill_color="#E6F3FF"
)
```

## Testing Strategy

### Unit Tests
- [ ] Table creation with various parameters
- [ ] Cell content setting and formatting
- [ ] Cell styling and color validation
- [ ] Range operations and boundary conditions
- [ ] Table structure modifications
- [ ] Error handling and edge cases

### Integration Tests
- [ ] End-to-end table creation and editing workflows
- [ ] Multiple tables on same slide
- [ ] Table operations with existing content
- [ ] Save/load preservation of table formatting
- [ ] Performance testing with large tables

### Edge Case Testing
- [ ] Maximum table dimensions (50x20)
- [ ] Empty cell handling
- [ ] Invalid coordinates and out-of-bounds operations
- [ ] Table deletion and cleanup
- [ ] Memory usage with multiple large tables

## Technical Considerations & Limitations

### python-pptx Library Limitations
- **Row/Column Insertion**: Limited support for dynamic row/column addition
- **Border Styling**: Complex border manipulation may require workarounds
- **Advanced Formatting**: Some PowerPoint table features not fully supported
- **Performance**: Large table operations may be slower

### Mitigation Strategies
- **Graceful Degradation**: Provide basic functionality where advanced features aren't available
- **Clear Documentation**: Document limitations and workarounds
- **Alternative Approaches**: Research alternative methods for unsupported operations
- **User Feedback**: Provide clear error messages for unsupported operations

## Future Enhancements

### Potential Phase 5+ Features
- [ ] Table templates and predefined styles
- [ ] Advanced border styling (individual sides)
- [ ] Cell merging and splitting
- [ ] Table-to-chart conversion
- [ ] Bulk data import from CSV/JSON
- [ ] Table sorting and filtering
- [ ] Formula support in cells
- [ ] Table accessibility features

### Integration Opportunities
- [ ] Integration with chart tools for data visualization
- [ ] Template-based table creation
- [ ] Table content search and replace
- [ ] Table export capabilities
- [ ] Collaborative editing features

## Success Metrics

### Functional Metrics
- [ ] All 6 table tools implemented and working
- [ ] 100% test coverage for table operations
- [ ] Zero breaking changes to existing functionality
- [ ] Performance benchmarks met for large tables

### User Experience Metrics
- [ ] Intuitive parameter naming and validation
- [ ] Clear and actionable error messages
- [ ] Comprehensive success feedback
- [ ] Consistent behavior with existing tools

### Quality Metrics
- [ ] Code review approval
- [ ] Documentation completeness
- [ ] Integration test suite passing
- [ ] Memory leak testing passed

## Risk Assessment

### High Risk Items
1. **python-pptx Limitations**: Some advanced features may not be implementable
   - *Mitigation*: Research and prototype early, document limitations
   
2. **Performance with Large Tables**: Memory and speed concerns
   - *Mitigation*: Implement performance testing, optimize critical paths
   
3. **Breaking Changes**: New features affecting existing functionality
   - *Mitigation*: Comprehensive regression testing, careful integration

### Medium Risk Items
1. **Complex Border Styling**: Advanced border features may be limited
2. **Error Handling**: Table-specific edge cases and validation
3. **Integration Complexity**: Ensuring tables work with all existing features

### Low Risk Items
1. **Basic Table Creation**: Well-supported by python-pptx
2. **Cell Content Management**: Straightforward implementation
3. **Success Message Integration**: Following existing patterns

## Conclusion

This implementation plan provides a comprehensive roadmap for adding robust table support to the PowerPoint MCP Server. The phased approach ensures stability while building towards full table functionality. The design maintains consistency with existing patterns while providing powerful new capabilities for table creation, editing, and styling.

The implementation will significantly enhance the server's capabilities, making it a more complete solution for PowerPoint automation and manipulation tasks. 