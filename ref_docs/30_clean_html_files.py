#!/usr/bin/env python3
"""
Script to clean HTML files by extracting only the <article> content.
This script processes all HTML files in functions/html/ directory and keeps only
the content within <article></article> tags, removing all other HTML.
"""

from bs4 import BeautifulSoup
import os
import sys
import shutil

def print_progress_bar(current, total, filename, bar_length=50):
    """Print a progress bar to show current progress."""
    percentage = (current / total) * 100
    filled_length = int(bar_length * current // total)
    bar = '█' * filled_length + '-' * (bar_length - filled_length)
    
    # Truncate filename if too long
    display_name = filename[:25] + "..." if len(filename) > 25 else filename
    
    sys.stdout.write(f'\r[{bar}] {percentage:.1f}% ({current}/{total}) Processing: {display_name}')
    sys.stdout.flush()
    
    if current == total:
        print()  # New line when complete

def extract_article_content(html_filepath):
    """Extract only the <article> content from an HTML file."""
    try:
        # Read the HTML file
        with open(html_filepath, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Parse with BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Find the article tag
        article = soup.find('article')
        
        if article:
            # Create a minimal HTML structure with just the article
            clean_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Function Documentation</title>
</head>
<body>
{article.prettify()}
</body>
</html>"""
            return clean_html
        else:
            # If no article tag found, return None
            return None
            
    except Exception as e:
        print(f"\nError processing {html_filepath}: {str(e)}")
        return None

def clean_all_html_files():
    """Process all HTML files and extract article content."""
    # Define directories
    html_dir = os.path.join('functions', 'html')
    backup_dir = os.path.join('functions', 'html_backup')
    
    # Check if HTML directory exists
    if not os.path.exists(html_dir):
        print("❌ Error: functions/html/ directory not found!")
        return
    
    # Get list of HTML files
    html_files = [f for f in os.listdir(html_dir) if f.endswith('.html')]
    
    if not html_files:
        print("❌ Error: No HTML files found in functions/html/ directory!")
        return
    
    print(f"🔧 HTML Article Extractor")
    print("=" * 50)
    print(f"📁 Processing {len(html_files)} HTML files")
    print(f"📍 Source: {html_dir}")
    print(f"💾 Backup: {backup_dir}")
    print("🎯 Extracting <article> content only...")
    
    # Ask for confirmation
    response = input(f"\n🚀 Ready to process {len(html_files)} HTML files? (y/N): ")
    
    if response.lower() not in ['y', 'yes']:
        print("\n🛑 Operation cancelled.")
        return
    
    # Create backup directory
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
        print(f"\n📂 Created backup directory: {backup_dir}")
    
    print("\n📊 Progress:")
    
    successful_extractions = 0
    failed_extractions = 0
    no_article_found = 0
    
    # Create a summary file
    summary_file = 'html_cleaning_summary.txt'
    
    with open(summary_file, 'w', encoding='utf-8') as summary:
        summary.write("HTML Article Extraction Summary\n")
        summary.write("=" * 50 + "\n\n")
        
        for i, filename in enumerate(html_files):
            # Update progress bar
            print_progress_bar(i + 1, len(html_files), filename)
            
            html_filepath = os.path.join(html_dir, filename)
            backup_filepath = os.path.join(backup_dir, filename)
            
            # Create backup of original file
            try:
                shutil.copy2(html_filepath, backup_filepath)
            except Exception as e:
                summary.write(f"❌ {filename}: Failed to create backup - {str(e)}\n")
                failed_extractions += 1
                continue
            
            # Extract article content
            clean_html = extract_article_content(html_filepath)
            
            if clean_html:
                try:
                    # Write the clean HTML back to the original file
                    with open(html_filepath, 'w', encoding='utf-8') as f:
                        f.write(clean_html)
                    
                    successful_extractions += 1
                    summary.write(f"✅ {filename}: Successfully extracted article content\n")
                    
                except Exception as e:
                    failed_extractions += 1
                    summary.write(f"❌ {filename}: Failed to write clean HTML - {str(e)}\n")
                    
            else:
                no_article_found += 1
                summary.write(f"⚠️  {filename}: No <article> tag found\n")
    
    # Final progress update
    print_progress_bar(len(html_files), len(html_files), "Complete!")
    
    # Calculate file size savings
    original_size = sum(os.path.getsize(os.path.join(backup_dir, f)) for f in os.listdir(backup_dir) if f.endswith('.html'))
    new_size = sum(os.path.getsize(os.path.join(html_dir, f)) for f in os.listdir(html_dir) if f.endswith('.html'))
    size_saved = original_size - new_size
    
    print(f"\n🎉 HTML cleaning complete!")
    print(f"📊 Results:")
    print(f"   ✅ Successfully processed: {successful_extractions}")
    print(f"   ❌ Failed extractions: {failed_extractions}")
    print(f"   ⚠️  No article tag found: {no_article_found}")
    print(f"   📁 Total files processed: {len(html_files)}")
    print(f"   📈 Success rate: {(successful_extractions/len(html_files))*100:.1f}%")
    print(f"   💾 Original size: {original_size:,} bytes")
    print(f"   📦 New size: {new_size:,} bytes")
    print(f"   💰 Space saved: {size_saved:,} bytes ({(size_saved/original_size)*100:.1f}%)")
    print(f"   📋 Summary saved to: {summary_file}")
    print(f"   🔙 Backups saved to: {backup_dir}")

if __name__ == "__main__":
    clean_all_html_files()
