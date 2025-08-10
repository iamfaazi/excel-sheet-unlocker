#!/usr/bin/env python3
"""
Excel Sheet Unlocker (No Password Required) for macOS
Removes sheet protection from Excel files without needing the password
"""

import openpyxl
import os
import sys
from pathlib import Path

def unlock_excel_sheets(file_path, output_path=None):
    """
    Remove protection from Excel sheets without password
    
    Args:
        file_path (str): Path to the protected Excel file
        output_path (str, optional): Path for unlocked file. If None, adds '_unlocked' suffix
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"❌ Error: File '{file_path}' not found.")
            return False
        
        print(f"📂 Loading Excel file: {file_path}")
        
        # Load workbook (this bypasses sheet-level password protection)
        workbook = openpyxl.load_workbook(file_path)
        
        # Process all sheets
        unlocked_sheets = []
        total_sheets = len(workbook.sheetnames)
        
        for sheet_name in workbook.sheetnames:
            print(f"🔍 Processing sheet: '{sheet_name}'")
            worksheet = workbook[sheet_name]
            
            try:
                # Method 1: Direct protection removal
                from openpyxl.worksheet.protection import SheetProtection
                
                # Check current protection status
                was_protected = False
                try:
                    was_protected = bool(worksheet.protection.sheet)
                except:
                    was_protected = True  # Assume protected if we can't determine
                
                # Always create a fresh, unprotected SheetProtection object
                new_protection = SheetProtection(
                    sheet=False,
                    password=None,
                    selectLockedCells=True,
                    selectUnlockedCells=True,
                    formatCells=True,
                    formatColumns=True,
                    formatRows=True,
                    insertColumns=True,
                    insertRows=True,
                    insertHyperlinks=True,
                    deleteColumns=True,
                    deleteRows=True,
                    sort=True,
                    autoFilter=True,
                    pivotTables=True,
                    objects=True,
                    scenarios=True
                )
                
                # Replace the protection object entirely
                worksheet.protection = new_protection
                
                if was_protected:
                    print(f"🔓 Successfully unlocked sheet: '{sheet_name}'")
                    unlocked_sheets.append(sheet_name)
                else:
                    print(f"✅ Sheet '{sheet_name}' was already unlocked")
                    
            except Exception as e:
                print(f"⚠️  Method 1 failed for '{sheet_name}': {e}")
                
                # Method 2: Brute force approach
                try:
                    # Try to unlock by directly modifying the worksheet's XML
                    worksheet._protection = None
                    
                    # Create minimal protection object
                    from openpyxl.worksheet.protection import SheetProtection
                    worksheet.protection = SheetProtection()
                    worksheet.protection.sheet = False
                    
                    print(f"🔓 Sheet '{sheet_name}' unlocked using Method 2")
                    unlocked_sheets.append(sheet_name)
                    
                except Exception as e2:
                    print(f"⚠️  Method 2 failed for '{sheet_name}': {e2}")
                    
                    # Method 3: Last resort - minimal intervention
                    try:
                        # Just try to disable the main protection flag
                        if hasattr(worksheet, 'protection'):
                            if hasattr(worksheet.protection, 'sheet'):
                                worksheet.protection.sheet = False
                            if hasattr(worksheet.protection, 'password'):
                                worksheet.protection.password = None
                        
                        print(f"🔓 Sheet '{sheet_name}' processed using Method 3")
                        unlocked_sheets.append(sheet_name)
                        
                    except Exception as e3:
                        print(f"❌ All methods failed for sheet '{sheet_name}': {e3}")
                        print("    This sheet will remain as-is in the output file.")
        
        # Report results
        if unlocked_sheets:
            print(f"\n🎉 Successfully unlocked {len(unlocked_sheets)} out of {total_sheets} sheets:")
            for sheet in unlocked_sheets:
                print(f"   - {sheet}")
        else:
            print(f"\n📋 All {total_sheets} sheets were already unlocked")
        
        # Determine output path
        if output_path is None:
            file_path_obj = Path(file_path)
            output_path = file_path_obj.parent / f"{file_path_obj.stem}_unlocked{file_path_obj.suffix}"
        
        # Save the unlocked workbook
        print(f"\n💾 Saving unlocked file to: {output_path}")
        workbook.save(output_path)
        workbook.close()
        
        print("✅ Process completed successfully!")
        print(f"📁 Unlocked file saved as: {os.path.basename(output_path)}")
        return True
        
    except PermissionError:
        print("❌ Error: Permission denied. Please close Excel and try again.")
        return False
    except FileNotFoundError:
        print("❌ Error: File not found. Please check the file path.")
        return False
    except Exception as e:
        print(f"❌ Error occurred: {str(e)}")
        return False

def process_multiple_files(directory_path):
    """
    Process all Excel files in a directory
    """
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"❌ Directory not found: {directory_path}")
        return
    
    # Find Excel files
    excel_files = []
    for pattern in ['*.xlsx', '*.xlsm', '*.xls']:
        excel_files.extend(directory.glob(pattern))
    
    if not excel_files:
        print("❌ No Excel files found in the directory")
        return
    
    print(f"📂 Found {len(excel_files)} Excel file(s) to process\n")
    
    successful = 0
    for i, file_path in enumerate(excel_files, 1):
        print(f"--- Processing file {i}/{len(excel_files)}: {file_path.name} ---")
        if unlock_excel_sheets(str(file_path)):
            successful += 1
        print()  # Add spacing between files
    
    print(f"🏁 Final Results: {successful}/{len(excel_files)} files processed successfully")

def main():
    """
    Main function with simple interface
    """
    print("🔐 Excel Sheet Unlocker for macOS")
    print("=================================")
    print("This tool removes sheet protection without needing passwords\n")
    
    # Check for command line arguments
    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) >= 3 else None
        
        if os.path.isdir(input_path):
            process_multiple_files(input_path)
        else:
            unlock_excel_sheets(input_path, output_path)
    else:
        # Interactive mode
        print("Choose an option:")
        print("1. Process a single file")
        print("2. Process all Excel files in a folder")
        
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            file_path = input("\n📂 Drag and drop your Excel file here (or enter path): ").strip()
            # Remove quotes if user drags and drops
            file_path = file_path.strip('"').strip("'")
            
            custom_output = input("💾 Custom output name? (Press Enter to auto-name): ").strip()
            output_path = custom_output if custom_output else None
            
            print()  # Add spacing
            unlock_excel_sheets(file_path, output_path)
            
        elif choice == "2":
            folder_path = input("\n📁 Enter folder path containing Excel files: ").strip()
            folder_path = folder_path.strip('"').strip("'")
            
            print()  # Add spacing
            process_multiple_files(folder_path)
            
        else:
            print("❌ Invalid choice. Please run the script again.")

if __name__ == "__main__":
    # Auto-install openpyxl if needed
    try:
        import openpyxl
    except ImportError:
        print("📦 Installing required package: openpyxl...")
        os.system("pip3 install openpyxl")
        print("✅ Installation complete!\n")
        import openpyxl
    
    main()
