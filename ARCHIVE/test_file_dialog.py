#!/usr/bin/env python3
"""
Test script to verify the file dialog functionality works correctly.
"""

"""
test_file_dialog.py

Smoke test for the tkinter file picker dialog in Insert-and-hide-a-column-improved.py.

Imports and calls select_file_with_dialog() directly, opens the file picker,
and reports whether a file was selected, its path, existence, and size.
Intended as a quick sanity check that tkinter is available and the dialog
renders correctly in the current environment before running the main script.

Usage:
    python test_file_dialog.py

    No arguments. Dialog opens immediately — select any file or cancel to exit.

Dependencies:
    - Insert-and-hide-a-column-improved.py must exist in the same directory
    - Requires tkinter (standard library) and a display environment (not headless)

Limitations:
    - Interactive only — cannot be run in CI or headless environments
    - Does not test any Excel operations, only the dialog UI component
    - Imports using underscore filename convention; ensure the source file has
      been renamed from 'Insert-and-hide-a-column-improved.py' to
      'Insert_and_hide_a_column_improved.py' or the import will fail
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from Insert_and_hide_a_column_improved import select_file_with_dialog

def test_file_dialog():
    """Test the file dialog function."""
    print("Testing file dialog functionality...")
    print("A file dialog should open. Please select a file or cancel.")
    print()
    
    try:
        file_path = select_file_with_dialog()
        if file_path:
            print(f"✅ File selected successfully: {file_path}")
            print(f"   File exists: {os.path.exists(file_path)}")
            print(f"   File size: {os.path.getsize(file_path)} bytes")
        else:
            print("ℹ️  No file selected (dialog cancelled)")
    except Exception as e:
        print(f"❌ Error testing file dialog: {e}")
        return False
    
    return True

if __name__ == "__main__":
    success = test_file_dialog()
    if success:
        print("\n✅ File dialog test completed successfully!")
    else:
        print("\n❌ File dialog test failed!")
        sys.exit(1)