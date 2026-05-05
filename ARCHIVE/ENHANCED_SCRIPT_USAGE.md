# Enhanced Copy-Insert-Hide Column Script

## Overview

The enhanced `copy-insert-hide-column.py` script automatically finds the last column containing a "COMMENTS DD-MM-YYYY" pattern and performs the following operations:

1. **Finds** the rightmost column with text matching "COMMENTS DD-MM-YYYY" format
2. **Copies** all data from that column
3. **Inserts** a new column before the target column
4. **Pastes** the copied data into the new column
5. **Hides** the newly inserted column

## Key Improvements

### ✨ Dynamic Pattern Matching
- **Before**: Hardcoded date "COMMENTS 19-03-2026"
- **After**: Automatically finds any date matching "COMMENTS DD-MM-YYYY" pattern

### 🔍 Last Column Detection
- Searches through all columns to find the **rightmost** column with the pattern
- Works with multiple COMMENTS columns in the same sheet

### 🛡️ Robust Error Handling
- Clear error messages when no matching column is found
- Graceful handling of edge cases
- Comprehensive logging of operations

### 📝 Better Code Structure
- Modular functions for maintainability
- Detailed documentation and comments
- Configurable constants for easy customization

## Usage

### Command Line
```bash
# Basic usage
python copy-insert-hide-column.py path/to/your/file.xlsx

# With backup creation
python copy-insert-hide-column.py path/to/your/file.xlsx --backup

# Interactive mode (prompts for file path)
python copy-insert-hide-column.py
```

### Supported Date Formats
The script recognizes these patterns:
- `COMMENTS 19-03-2026`
- `COMMENTS  15-04-2026` (extra spaces)
- `comments 10-05-2026` (case insensitive)
- Any date in DD-MM-YYYY format

### Example Workflow

**Before Script:**
```
| A | B | C | D | E |
|---|---|---|---|---|
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   | COMMENTS 25-03-2026 |
|   |   |   |   | Comment data |
|   |   |   |   | More data |
```

**After Script:**
```
| A | B | C | D | [E Hidden] | F |
|---|---|---|---|------------|---|
|   |   |   |   |   (hidden) |   |
|   |   |   |   |   (hidden) |   |
|   |   |   |   |   (hidden) |   |
|   |   |   |   |   (hidden) |   |
|   |   |   |   |   (hidden) |   |
|   |   |   |   |   (hidden) | COMMENTS 25-03-2026 |
|   |   |   |   |   (hidden) | Comment data |
|   |   |   |   |   (hidden) | More data |
```

## Configuration

### Constants (can be modified in the script)
```python
COMMENTS_PATTERN = r'COMMENTS\s+\d{2}-\d{2}-\d{4}'  # Regex pattern to match
SEARCH_ROW = 6                                      # Row to search for pattern
```

### Customization Options
- Change `COMMENTS_PATTERN` to match different text formats
- Modify `SEARCH_ROW` to search in a different row
- Adjust the search logic for different requirements

## Error Handling

The script provides clear feedback for common issues:

- **No matching column found**: "No column found with pattern 'COMMENTS\s+\d{2}-\d{2}-\d{4}' in row 6"
- **File not found**: "File not found: [path]"
- **Processing errors**: Detailed error messages with context

## Dependencies

- `xlwings`: Excel automation
- `argparse`: Command line argument parsing
- `re`: Regular expressions
- `pathlib`: File path handling
- `shutil`: File operations

## Testing

Run the included test script to verify functionality:
```bash
python test_enhanced_script.py
```

This creates a test Excel file with multiple COMMENTS columns and verifies that the script correctly identifies and processes the last one.

## Benefits

1. **No manual date updates** - Works with any date format
2. **Automatic detection** - Finds the correct column every time
3. **Reusable** - Works across different files and dates
4. **Safe** - Optional backup creation
5. **Transparent** - Clear logging of all operations