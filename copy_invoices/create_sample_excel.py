#!/usr/bin/env python3
"""
Simple test/demo script for the OneDrive downloader.
This creates a sample Excel file for testing purposes.
"""

from pathlib import Path
from openpyxl import Workbook

def create_sample_excel():
    """Create a sample Excel file for testing."""
    excel_path = Path("sample_data.xlsx")
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Downloads"
    
    # Add headers
    ws['A1'] = "OneDrive Link"
    ws['B1'] = "Local Path"
    ws['C1'] = "Status"
    
    # Make headers bold
    for col in ['A1', 'B1', 'C1']:
        ws[col].font = ws[col].font.copy()
    
    # Add sample data (replace with actual OneDrive links)
    sample_links = [
        "https://1drv.ms/b/c/e7cee1d47fca9dd4/IQDiMORSPfCCQ6IzNV_HwfGhAc4-rb87dHMKwk-9-_G2jI0?e=bXz2c9",
        "https://1drv.ms/x/c/YOUR_ACTUAL_LINK_HERE",
        "https://1drv.ms/w/c/YOUR_ACTUAL_LINK_HERE",
    ]
    
    for idx, link in enumerate(sample_links, start=2):
        ws[f'A{idx}'] = link
        ws[f'C{idx}'] = "Pending"
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 80
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 15
    
    # Save
    wb.save(excel_path)
    print(f"✅ Sample Excel file created: {excel_path}")
    print(f"\nPlease replace the sample OneDrive links in column A with your actual links.")
    print(f"\nThen run the downloader with:")
    print(f"  python download_onedrive_files.py -e sample_data.xlsx -s A -t B -p ./downloads")


if __name__ == '__main__':
    create_sample_excel()
