#!/usr/bin/env python3
"""
Script to download Microsoft Excel function documentation HTML pages.
This script downloads the raw HTML files and stores them in the functions folder.
"""

import requests
from bs4 import BeautifulSoup
import re
import time
import os
import sys
from urllib.parse import urljoin

BASE_URL = "https://support.microsoft.com/"

def print_progress_bar(current, total, function_name, bar_length=50):
    """Print a progress bar to show current progress."""
    percentage = (current / total) * 100
    filled_length = int(bar_length * current // total)
    bar = '█' * filled_length + '-' * (bar_length - filled_length)
    
    # Truncate function name if too long
    display_name = function_name[:20] + "..." if len(function_name) > 20 else function_name
    
    sys.stdout.write(f'\r[{bar}] {percentage:.1f}% ({current}/{total}) Downloading: {display_name}')
    sys.stdout.flush()
    
    if current == total:
        print()  # New line when complete

def create_function_filename(function_name):
    """Create a safe filename from function name."""
    # Remove version info and parentheses
    clean_name = re.sub(r'\s*\(\d{4}\)', '', function_name)
    # Replace special characters
    safe_name = re.sub(r'[^\w\s.-]', '_', clean_name.strip())
    # Replace spaces and multiple underscores
    safe_name = re.sub(r'[\s_]+', '_', safe_name)
    # Convert to lowercase
    return safe_name.lower()

def download_function_html(url, function_name, output_dir):
    """Download the HTML content for a function and save it."""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # Create filename
        safe_filename = create_function_filename(function_name)
        html_filepath = os.path.join(output_dir, f"{safe_filename}.html")
        
        # Save the HTML content
        with open(html_filepath, 'w', encoding='utf-8') as f:
            f.write(response.text)
        
        return True, html_filepath
        
    except Exception as e:
        return False, str(e)

def download_all_function_docs():
    """Download HTML documentation for all functions from the original HTML file."""
    # Check if download summary exists to resume failed downloads
    summary_file = 'download_summary.txt'
    failed_functions = []
    existing_summary = {}
    
    if os.path.exists(summary_file):
        print("📋 Found existing download summary. Analyzing failed downloads...")
        
        # Read existing summary to find failed downloads
        with open(summary_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for line in lines:
            line = line.strip()
            if line.startswith('❌'):
                # Extract function name from failed entry
                # Format: "❌ FUNCTION_NAME: error message"
                # Find the first colon that's followed by a space and error message
                colon_pos = line.find(': ', 2)  # Start search after "❌ "
                if colon_pos != -1:
                    function_name = line[2:colon_pos].strip()  # Remove "❌ " prefix
                    failed_functions.append(function_name)
            elif line.startswith('✅'):
                # Track successful downloads
                # Format: "✅ FUNCTION_NAME: file_path"
                colon_pos = line.find(': ', 2)  # Start search after "✅ "
                if colon_pos != -1:
                    function_name = line[2:colon_pos].strip()  # Remove "✅ " prefix
                    existing_summary[function_name] = line
        
        print(f"📊 Found {len(failed_functions)} failed downloads to retry")
        
        if len(failed_functions) > 0:
            print(f"🔍 Failed functions to retry: {failed_functions[:5]}{'...' if len(failed_functions) > 5 else ''}")
        
        if not failed_functions:
            print("🎉 All downloads were successful! Nothing to retry.")
            return
    
    # Read the HTML file
    try:
        with open('functions.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        print("❌ Error: functions.html not found in current directory!")
        return
    
    # Parse the HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find all table rows (skip header)
    table = soup.find('table')
    if not table:
        print("❌ Error: No table found in HTML file!")
        return
    
    rows = table.find_all('tr')[1:]  # Skip header row
    
    # Filter rows based on whether we're resuming or starting fresh
    if failed_functions:
        # Only process failed functions
        rows_to_process = []
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 2:
                function_name = cells[0].get_text().strip()
                # Check if this function matches any of the failed functions
                # Handle exact match or partial match (accounting for version differences)
                for failed_func in failed_functions:
                    if (function_name == failed_func or 
                        function_name.split()[0] == failed_func.split()[0]):  # Match base function name
                        rows_to_process.append(row)
                        break
        
        print(f"🔄 Resuming downloads for {len(rows_to_process)} failed functions...")
    else:
        # Process all functions (first time run)
        rows_to_process = rows
        print(f"🚀 Starting fresh download of {len(rows_to_process)} Excel functions...")
    
    total_functions = len(rows_to_process)
    
    if total_functions == 0:
        print("❌ No functions to process!")
        return
    
    # Create output directory
    html_output_dir = os.path.join('functions', 'html')
    os.makedirs(html_output_dir, exist_ok=True)
    
    print(f"⏱️  With 1-second delays, this will take approximately {total_functions//60} minutes")
    print(f"📁 HTML files will be saved to: {html_output_dir}")
    print("📊 Progress:")
    
    successful_downloads = 0
    failed_downloads = 0
    start_time = time.time()
    
    # Prepare summary update
    if failed_functions:
        # Load existing summary for updates
        summary_data = existing_summary.copy()
    else:
        # Start fresh summary
        summary_data = {}
    
    for i, row in enumerate(rows_to_process):
        cells = row.find_all('td')
        if len(cells) >= 2:
            # Extract function info
            function_cell = cells[0]
            description_cell = cells[1]
            
            # Get function name (remove any HTML tags)
            function_name = function_cell.get_text().strip()
            
            # Update progress bar
            print_progress_bar(i + 1, total_functions, function_name)
            
            # Get description and extract type/description
            full_desc = description_cell.get_text().strip()
            desc_parts = full_desc.split(': ', 1)
            if len(desc_parts) == 2:
                function_type = desc_parts[0]
                description = desc_parts[1]
            else:
                function_type = "General"
                description = full_desc
            
            # Get URL from the link
            link = function_cell.find('a')
            if link and link.get('href'):
                url = urljoin(BASE_URL, link.get('href'))
            else:
                failed_downloads += 1
                summary_data[function_name] = f"❌ {function_name}: No URL found"
                continue
            
            # Add delay to be respectful to the server (except for first request)
            if i > 0:
                time.sleep(1)  # 1 second delay between requests
            
            # Download the HTML
            success, result = download_function_html(url, function_name, html_output_dir)
            
            if success:
                successful_downloads += 1
                summary_data[function_name] = f"✅ {function_name}: {result}"
            else:
                failed_downloads += 1
                summary_data[function_name] = f"❌ {function_name}: {result}"
    
    # Final progress update
    print_progress_bar(total_functions, total_functions, "Complete!")
    
    # Write updated summary file
    with open(summary_file, 'w', encoding='utf-8') as f:
        f.write("Excel Function HTML Download Summary\n")
        f.write("=" * 50 + "\n\n")
        
        # Write all entries (existing + new/updated) in sorted order
        for function_name in sorted(summary_data.keys()):
            f.write(summary_data[function_name] + "\n")
    
    # Calculate elapsed time
    elapsed_time = time.time() - start_time
    minutes = int(elapsed_time // 60)
    seconds = int(elapsed_time % 60)
    
    total_successful = len([v for v in summary_data.values() if v.startswith('✅')])
    total_failed = len([v for v in summary_data.values() if v.startswith('❌')])
    total_processed = len(summary_data)
    
    print(f"\n🎉 Download complete!")
    print(f"⏱️  Total time: {minutes}m {seconds}s")
    print(f"📊 Overall Results:")
    print(f"   ✅ Total successful downloads: {total_successful}")
    print(f"   ❌ Total failed downloads: {total_failed}")
    print(f"   📁 Total functions in summary: {total_processed}")
    print(f"   📈 Overall success rate: {(total_successful/total_processed)*100:.1f}%")
    
    if failed_functions:
        retry_successful = len([v for k, v in summary_data.items() if k in failed_functions and v.startswith('✅')])
        print(f"   🔄 Retry session results:")
        print(f"      ✅ Newly successful: {retry_successful}")
        print(f"      ❌ Still failing: {total_functions - retry_successful}")
    
    print(f"   📋 Updated summary saved to: {summary_file}")

if __name__ == "__main__":
    print("🔧 Excel Function HTML Downloader")
    print("=" * 50)
    
    # Check if functions.html exists
    if not os.path.exists('functions.html'):
        print("❌ Error: functions.html not found in current directory!")
        sys.exit(1)
    
    # Check if this is a resume operation or fresh start
    summary_file = 'download_summary.txt'
    if os.path.exists(summary_file):
        print("\n📋 Found existing download summary.")
        print("🔄 This will retry only the failed downloads from previous run.")
        print("📁 Successful downloads will be preserved.")
        print("🌐 Failed downloads will be retried with fresh requests.")
    else:
        print("\n📋 No existing summary found - starting fresh download.")
        print("🌐 This will download HTML pages for all 515+ Excel functions.")
        print("⏱️  With 1-second delays between requests, this will take approximately 8-10 minutes.")
        print("📁 HTML files will be saved in functions/html/ directory.")
    
    response = input("\n🚀 Ready to start/resume downloading? (y/N): ")
    
    if response.lower() in ['y', 'yes']:
        print(f"\n🔄 {'Resuming' if os.path.exists(summary_file) else 'Starting'} HTML download process...")
        download_all_function_docs()
        print("\n✨ All done! HTML files are ready for processing.")
    else:
        print("\n🛑 Operation cancelled.")
        print("💡 Tip: Run this script to download/retry HTML files, then use other scripts for processing.")
