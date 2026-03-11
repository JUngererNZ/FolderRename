#!/usr/bin/env python3
"""
Test script to verify the file dialog functionality works correctly.
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