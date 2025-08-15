#!/usr/bin/env python3
"""
Extract only <article> content from HTML files.

This script processes all HTML files in the functions/html/ directory
and extracts only the content within <article> tags, removing all
other HTML structure (head, body, etc.) for easier integration.
"""

import os
import sys
from pathlib import Path
from bs4 import BeautifulSoup
import time

def extract_article_content(html_file_path):
    """
    Extract only the <article> content from an HTML file.
    
    Args:
        html_file_path (str): Path to the HTML file
        
    Returns:
        str: The extracted article content as HTML string, or None if no article found
    """
    try:
        with open(html_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        soup = BeautifulSoup(content, 'html.parser')
        
        # Find the article tag
        article = soup.find('article')
        
        if article:
            # Return the article content as pretty-printed HTML
            return str(article.prettify())
        else:
            print(f"⚠️  No <article> tag found in {html_file_path}")
            return None
            
    except Exception as e:
        print(f"❌ Error processing {html_file_path}: {e}")
        return None

def process_all_html_files():
    """
    Process all HTML files in the functions/html/ directory.
    """
    html_dir = Path("functions/html")
    
    if not html_dir.exists():
        print(f"❌ Directory {html_dir} does not exist!")
        return
    
    # Get all HTML files
    html_files = list(html_dir.glob("*.html"))
    total_files = len(html_files)
    
    if total_files == 0:
        print("❌ No HTML files found in functions/html/")
        return
    
    print(f"🔍 Found {total_files} HTML files to process")
    print("🔄 Extracting <article> content from HTML files...")
    
    # Create backup directory
    backup_dir = Path("functions/html_backup")
    if not backup_dir.exists():
        backup_dir.mkdir(parents=True)
        print(f"📁 Created backup directory: {backup_dir}")
    
    processed = 0
    errors = 0
    space_saved = 0
    
    for html_file in html_files:
        try:
            # Get original file size
            original_size = html_file.stat().st_size
            
            # Create backup if it doesn't exist
            backup_file = backup_dir / html_file.name
            if not backup_file.exists():
                import shutil
                shutil.copy2(html_file, backup_file)
            
            # Extract article content
            article_content = extract_article_content(html_file)
            
            if article_content:
                # Write back only the article content
                with open(html_file, 'w', encoding='utf-8') as file:
                    file.write(article_content)
                
                # Calculate space saved
                new_size = html_file.stat().st_size
                space_saved += (original_size - new_size)
                
                processed += 1
                
                # Progress indicator
                if processed % 50 == 0 or processed == total_files:
                    progress = (processed / total_files) * 100
                    print(f"📈 Progress: {processed}/{total_files} files ({progress:.1f}%)")
            else:
                errors += 1
                
        except Exception as e:
            print(f"❌ Error processing {html_file.name}: {e}")
            errors += 1
    
    # Summary
    print("\n" + "="*60)
    print("📊 EXTRACTION SUMMARY")
    print("="*60)
    print(f"✅ Files processed successfully: {processed}")
    print(f"❌ Files with errors: {errors}")
    print(f"💾 Space saved: {space_saved / (1024*1024):.1f} MB")
    print(f"📁 Original files backed up to: {backup_dir}")
    print("="*60)

def main():
    """Main function."""
    print("🚀 Starting article content extraction...")
    print("📝 This will extract only <article> content from HTML files")
    print("💾 Original files will be backed up automatically")
    print()
    
    start_time = time.time()
    process_all_html_files()
    end_time = time.time()
    
    print(f"⏱️  Total execution time: {end_time - start_time:.2f} seconds")
    print("✅ Article extraction complete!")

if __name__ == "__main__":
    main()
