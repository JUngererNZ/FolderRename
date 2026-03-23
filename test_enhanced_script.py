#!/usr/bin/env python3
"""
Test script to validate the enhanced copy-insert-hide-column.py functionality.
This script creates a test Excel file and runs the enhanced script on it.
"""

import xlwings as xw
import tempfile
import os
from pathlib import Path
import subprocess
import sys

def create_test_excel():
    """Create a test Excel file with multiple COMMENTS columns"""
    # Create a temporary Excel file
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_path = Path(temp_file.name)
    temp_file.close()
    
    # Create test workbook
    wb = xw.Book()
    ws = wb.sheets[0]
    ws.name = "Test Sheet"
    
    # Create test data with multiple COMMENTS columns
    test_data = [
        ["Data 1", "Data 2", "Data 3", "Data 4", "Data 5"],
        ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3", "Row 2 Col 4", "Row 2 Col 5"],
        ["Row 3 Col 1", "Row 3 Col 2", "Row 3 Col 3", "Row 3 Col 4", "Row 3 Col 5"],
        ["Row 4 Col 1", "Row 4 Col 2", "Row 4 Col 3", "Row 4 Col 4", "Row 4 Col 5"],
        ["Row 5 Col 1", "Row 5 Col 2", "Row 5 Col 3", "Row 5 Col 4", "Row 5 Col 5"],
        ["COMMENTS 10-03-2026", "COMMENTS 15-03-2026", "COMMENTS 20-03-2026", "COMMENTS 25-03-2026", "COMMENTS 30-03-2026"],
        ["Comment 1", "Comment 2", "Comment 3", "Comment 4", "Comment 5"],
        ["More data 1", "More data 2", "More data 3", "More data 4", "More data 5"]
    ]
    
    # Fill the worksheet with test data
    for row_idx, row_data in enumerate(test_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.range((row_idx, col_idx)).value = value
    
    # Save and close
    wb.save(str(temp_path))
    wb.close()
    
    return temp_path

def run_enhanced_script(test_file):
    """Run the enhanced script on the test file"""
    script_path = Path(__file__).parent / "copy-insert-hide-column.py"
    
    try:
        # Run the script with the test file
        result = subprocess.run([
            sys.executable, str(script_path), str(test_file), "--backup"
        ], capture_output=True, text=True, timeout=30)
        
        print("=== SCRIPT OUTPUT ===")
        print(result.stdout)
        if result.stderr:
            print("=== SCRIPT ERRORS ===")
            print(result.stderr)
        
        return result.returncode == 0
        
    except subprocess.TimeoutExpired:
        print("❌ Script execution timed out")
        return False
    except Exception as e:
        print(f"❌ Error running script: {e}")
        return False

def verify_results(test_file):
    """Verify that the script worked correctly"""
    try:
        wb = xw.Book(str(test_file))
        ws = wb.sheets[0]
        
        print("\n=== VERIFICATION RESULTS ===")
        
        # Check if the last COMMENTS column was found (should be column E with "COMMENTS 30-03-2026")
        expected_column = 5  # Column E
        expected_value = "COMMENTS 30-03-2026"
        
        # Check if a new column was inserted before column E (should now be column E, original E becomes F)
        # And check if the new column is hidden
        new_column_hidden = ws.range((1, expected_column)).api.EntireColumn.Hidden
        original_column_value = ws.range((6, expected_column + 1)).value  # Should still be "COMMENTS 30-03-2026"
        
        print(f"Expected target column: E (column {expected_column})")
        print(f"Expected target value: {expected_value}")
        print(f"New column (E) is hidden: {new_column_hidden}")
        print(f"Original column (F) still contains: '{original_column_value}'")
        
        # Check if data was copied correctly
        original_data = ws.range((1, expected_column + 1), (8, expected_column + 1)).value
        copied_data = ws.range((1, expected_column), (8, expected_column)).value
        
        data_matches = original_data == copied_data
        print(f"Data copied correctly: {data_matches}")
        
        success = new_column_hidden and original_column_value == expected_value and data_matches
        
        wb.close()
        return success
        
    except Exception as e:
        print(f"❌ Error verifying results: {e}")
        if 'wb' in locals():
            wb.close()
        return False

def main():
    """Main test function"""
    print("🧪 Testing Enhanced Copy-Insert-Hide Column Script")
    print("=" * 50)
    
    # Create test file
    print("1. Creating test Excel file...")
    test_file = create_test_excel()
    print(f"   Test file created: {test_file}")
    
    try:
        # Run the enhanced script
        print("\n2. Running enhanced script...")
        script_success = run_enhanced_script(test_file)
        
        if script_success:
            print("   ✅ Script executed successfully")
            
            # Verify results
            print("\n3. Verifying results...")
            verification_success = verify_results(test_file)
            
            if verification_success:
                print("   ✅ All verifications passed!")
                print("\n🎉 TEST PASSED: Enhanced script works correctly!")
            else:
                print("   ❌ Verification failed")
                print("\n❌ TEST FAILED: Script didn't work as expected")
        else:
            print("   ❌ Script execution failed")
            print("\n❌ TEST FAILED: Script couldn't run")
            
    finally:
        # Clean up
        if test_file.exists():
            test_file.unlink()
            print(f"\n🧹 Cleaned up test file: {test_file}")

if __name__ == "__main__":
    main()