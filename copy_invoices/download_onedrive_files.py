#!/usr/bin/env python3
"""
OneDrive File Downloader Script

Downloads files from OneDrive links stored in an Excel file and saves them to a local path.
Maintains the folder structure and updates the Excel file with local paths.
"""

import argparse
import logging
import sys
from pathlib import Path
from urllib.parse import urlparse
import shutil
from datetime import datetime
from typing import Dict, List, Tuple
import json

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


class OneDriveDownloader:
    """Handles downloading files from OneDrive and managing Excel updates."""
    
    def __init__(self, excel_path: str, excel_column_source: str, 
                 excel_column_target: str, target_path: str, log_file: str = None):
        """
        Initialize the downloader.
        
        Args:
            excel_path: Path to the Excel file
            excel_column_source: Column letter/name for source OneDrive links (e.g., 'A' or 'Link')
            excel_column_target: Column letter/name for target local paths (e.g., 'B' or 'LocalPath')
            target_path: Local target directory for downloads
            log_file: Optional custom log file path
        """
        self.excel_path = Path(excel_path)
        self.excel_column_source = excel_column_source
        self.excel_column_target = excel_column_target
        self.target_path = Path(target_path)
        
        # Create target directory if it doesn't exist
        self.target_path.mkdir(parents=True, exist_ok=True)
        
        # Setup logging
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if log_file is None:
            log_file = self.target_path / f"download_report_{timestamp}.log"
        else:
            log_file = Path(log_file)
        
        self.log_file = log_file
        self.setup_logging()
        
        # Statistics
        self.stats = {
            "total": 0,
            "successful": 0,
            "failed": 0,
            "skipped": 0,
            "errors": []
        }
    
    def setup_logging(self):
        """Configure logging with both file and console output."""
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.DEBUG)
        
        # Remove existing handlers
        logger.handlers.clear()
        
        # File handler
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        
        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        self.logger = logger
    
    def convert_onedrive_link(self, link: str) -> str:
        """
        Convert OneDrive share link to direct download link.
        
        Args:
            link: OneDrive link (e.g., https://1drv.ms/b/c/...)
            
        Returns:
            Direct download URL
        """
        try:
            # OneDrive links end with ?e=XXXX, we need to remove that and add &download=1
            if '?' in link:
                base_link = link.split('?')[0]
            else:
                base_link = link
            
            # Convert to download URL
            download_link = base_link.rstrip('/') + '?download=1'
            return download_link
        except Exception as e:
            self.logger.error(f"Error converting OneDrive link: {e}")
            return None
    
    def extract_filename_from_url(self, url: str) -> str:
        """
        Extract filename from URL.
        
        Args:
            url: The URL
            
        Returns:
            Filename or None
        """
        try:
            path = urlparse(url).path
            return path.split('/')[-1] if path else None
        except Exception as e:
            self.logger.error(f"Error extracting filename from URL: {e}")
            return None
    
    def download_file(self, onedrive_link: str, local_path: Path, timeout: int = 30) -> Tuple[bool, str]:
        """
        Download a file from OneDrive link.
        
        Args:
            onedrive_link: OneDrive link
            local_path: Local path where file should be saved
            timeout: Request timeout in seconds
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Convert to download URL
            download_url = self.convert_onedrive_link(onedrive_link)
            if not download_url:
                return False, "Failed to convert OneDrive link"
            
            # Create parent directories
            local_path.parent.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Downloading from: {onedrive_link}")
            
            # Download file with timeout and retries
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(download_url, timeout=timeout, headers=headers, 
                                   stream=True, allow_redirects=True)
            response.raise_for_status()
            
            # Save file
            with open(local_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            file_size = local_path.stat().st_size
            self.logger.info(f"Successfully downloaded: {local_path.name} ({file_size} bytes)")
            return True, str(local_path)
            
        except requests.exceptions.Timeout:
            msg = f"Timeout downloading file (timeout={timeout}s): {onedrive_link}"
            self.logger.error(msg)
            return False, msg
        except requests.exceptions.ConnectionError as e:
            msg = f"Connection error: {str(e)}"
            self.logger.error(msg)
            return False, msg
        except requests.exceptions.HTTPError as e:
            msg = f"HTTP error {e.response.status_code}: {str(e)}"
            self.logger.error(msg)
            return False, msg
        except Exception as e:
            msg = f"Unexpected error downloading file: {str(e)}"
            self.logger.error(msg)
            return False, msg
    
    def get_column_index(self, column_identifier: str) -> int:
        """
        Convert column identifier (letter or name) to index.
        
        Args:
            column_identifier: Column letter (e.g., 'A', 'B') or column name
            
        Returns:
            Column index (1-based)
        """
        try:
            # Try as column letter first
            if len(column_identifier) <= 2 and column_identifier.isalpha():
                from openpyxl.utils import column_index_from_string
                return column_index_from_string(column_identifier)
            else:
                # Try as integer
                return int(column_identifier)
        except Exception:
            self.logger.error(f"Invalid column identifier: {column_identifier}")
            return None
    
    def process_excel(self) -> bool:
        """
        Process Excel file: download files and update paths.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load workbook
            if not self.excel_path.exists():
                self.logger.error(f"Excel file not found: {self.excel_path}")
                return False
            
            workbook = load_workbook(self.excel_path)
            worksheet = workbook.active
            
            self.logger.info(f"Processing Excel file: {self.excel_path}")
            self.logger.info(f"Source column: {self.excel_column_source}, Target column: {self.excel_column_target}")
            
            # Get column indices
            source_col = self.get_column_index(self.excel_column_source)
            target_col = self.get_column_index(self.excel_column_target)
            
            if source_col is None or target_col is None:
                self.logger.error("Invalid column identifiers")
                return False
            
            self.logger.info(f"Source column index: {source_col}, Target column index: {target_col}")
            
            # Iterate through rows
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):  # Skip header
                try:
                    source_cell = row[source_col - 1]
                    target_cell = row[target_col - 1]
                    
                    # Try to get hyperlink URL first, fallback to cell value
                    if source_cell.hyperlink:
                        onedrive_link = source_cell.hyperlink.target
                    else:
                        onedrive_link = source_cell.value
                    
                    # Skip empty cells
                    if not onedrive_link or not isinstance(onedrive_link, str):
                        self.logger.debug(f"Row {row_idx}: Skipping empty cell")
                        self.stats["skipped"] += 1
                        continue
                    
                    onedrive_link = onedrive_link.strip()
                    self.stats["total"] += 1
                    
                    self.logger.info(f"\nRow {row_idx}: Processing {onedrive_link}")
                    
                    # Create local path maintaining structure
                    # Extract filename from URL
                    filename = self.extract_filename_from_url(onedrive_link)
                    if not filename:
                        filename = f"file_{row_idx}"
                    
                    local_file_path = self.target_path / filename
                    
                    # Download file
                    success, message = self.download_file(onedrive_link, local_file_path)
                    
                    if success:
                        # Update Excel with local path
                        target_cell.value = message
                        self.stats["successful"] += 1
                        self.logger.info(f"Row {row_idx}: Updated with local path: {message}")
                    else:
                        self.stats["failed"] += 1
                        self.stats["errors"].append({
                            "row": row_idx,
                            "link": onedrive_link,
                            "error": message
                        })
                        self.logger.error(f"Row {row_idx}: {message}")
                        # Leave target cell empty or keep previous value
                        
                except Exception as e:
                    self.stats["failed"] += 1
                    error_msg = f"Error processing row {row_idx}: {str(e)}"
                    self.logger.error(error_msg)
                    self.stats["errors"].append({
                        "row": row_idx,
                        "error": error_msg
                    })
                    # Continue with next row
                    continue
            
            # Save updated workbook
            self.logger.info("\nSaving updated Excel file...")
            workbook.save(self.excel_path)
            self.logger.info("Excel file saved successfully")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Fatal error processing Excel file: {str(e)}")
            return False
    
    def generate_report(self):
        """Generate and display the final report."""
        report = f"""
{'='*60}
DOWNLOAD REPORT
{'='*60}
Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Excel File: {self.excel_path}
Target Path: {self.target_path}
Log File: {self.log_file}

STATISTICS:
-----------
Total Files: {self.stats['total']}
Successful: {self.stats['successful']}
Failed: {self.stats['failed']}
Skipped: {self.stats['skipped']}

Success Rate: {(self.stats['successful'] / max(self.stats['total'], 1) * 100):.1f}%

"""
        
        if self.stats['errors']:
            report += f"\nFAILED DOWNLOADS ({len(self.stats['errors'])}):\n"
            report += "-" * 60 + "\n"
            for error in self.stats['errors']:
                report += f"Row {error.get('row', 'N/A')}: {error.get('link', 'N/A')}\n"
                report += f"  Error: {error.get('error', 'Unknown error')}\n\n"
        
        report += f"{'='*60}\n"
        
        # Print report
        print(report)
        self.logger.info(report)
        
        # Save report to JSON file
        json_report_path = self.target_path / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(json_report_path, 'w', encoding='utf-8') as f:
            json.dump({
                "timestamp": datetime.now().isoformat(),
                "excel_file": str(self.excel_path),
                "target_path": str(self.target_path),
                "log_file": str(self.log_file),
                "statistics": self.stats
            }, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"JSON report saved: {json_report_path}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Download files from OneDrive links in Excel and update local paths',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python download_onedrive_files.py -e data.xlsx -s A -t B -p ./downloads
  python download_onedrive_files.py --excel data.xlsx --source-col A --target-col B --target-path ./downloads
        """
    )
    
    parser.add_argument('-e', '--excel', dest='excel_path',
                       required=True,
                       help='Path to Excel file')
    parser.add_argument('-s', '--source-col', dest='excel_column_source',
                       required=True,
                       help='Source column (column letter like A, B or column index like 1, 2)')
    parser.add_argument('-t', '--target-col', dest='excel_column_target',
                       required=True,
                       help='Target column (column letter like A, B or column index like 1, 2)')
    parser.add_argument('-p', '--target-path', dest='target_path',
                       required=True,
                       help='Local target path for downloads')
    parser.add_argument('-l', '--log-file', dest='log_file',
                       help='Optional custom log file path')
    
    args = parser.parse_args()
    
    # Create downloader and process
    downloader = OneDriveDownloader(
        excel_path=args.excel_path,
        excel_column_source=args.excel_column_source,
        excel_column_target=args.excel_column_target,
        target_path=args.target_path,
        log_file=args.log_file
    )
    
    downloader.logger.info("Starting OneDrive file download process")
    
    success = downloader.process_excel()
    
    downloader.generate_report()
    
    if success and downloader.stats['failed'] == 0:
        downloader.logger.info("Process completed successfully!")
        return 0
    elif success:
        downloader.logger.warning(f"Process completed with {downloader.stats['failed']} failures")
        return 1
    else:
        downloader.logger.error("Process failed!")
        return 2


if __name__ == '__main__':
    sys.exit(main())
