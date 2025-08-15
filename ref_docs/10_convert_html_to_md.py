#!/usr/bin/env python3
import re
import html
import os

BASE_URL = "https://support.microsoft.com/"

def create_function_filename(func_name):
    """Create a safe filename from function name"""
    # Remove version info for filename
    clean_name = re.sub(r'\s*\([^)]+\)\s*', '', func_name)
    # Replace special characters with underscores
    safe_name = re.sub(r'[^\w\s-]', '_', clean_name)
    # Replace spaces with underscores and convert to lowercase
    safe_name = re.sub(r'\s+', '_', safe_name).lower()
    # Remove multiple underscores
    safe_name = re.sub(r'_+', '_', safe_name)
    # Remove leading/trailing underscores
    safe_name = safe_name.strip('_')
    return f"{safe_name}.md"

def create_individual_function_file(func_name, func_type, desc_text, original_url):
    """Create individual markdown file for each function"""
    filename = create_function_filename(func_name)
    filepath = os.path.join('functions', filename)
    
    # Create the functions directory if it doesn't exist
    os.makedirs('functions', exist_ok=True)
    
    # Create the content for individual function file
    content = f"""# {func_name}

**Category:** {func_type}

## Description
{desc_text}

## Microsoft Documentation
[Official Documentation]({BASE_URL}{original_url})

---
*This function reference was automatically generated.*
"""
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return filename

# Read the HTML file
with open('functions.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Parse table rows
rows = re.findall(r'<tr>(.*?)</tr>', content, re.DOTALL)

markdown_lines = []
markdown_lines.append('# Excel Functions Reference')
markdown_lines.append('')
markdown_lines.append('| S.No | Function Name | Type | Description |')
markdown_lines.append('|------|---------------|------|-------------|')

function_count = 0

# Skip the header row (first row)
for row in rows[1:]:  # Skip first row which is the header
    # Extract cell data
    cells = re.findall(r'<td>(.*?)</td>', row, re.DOTALL)
    
    if len(cells) >= 2:
        # Extract function name from first cell
        func_name_match = re.search(r'<a[^>]*>([^<]+)</a>', cells[0])
        original_url = ""
        
        if func_name_match:
            func_name = func_name_match.group(1).strip()
            
            # Extract original URL
            url_match = re.search(r'href="([^"]+)"', cells[0])
            if url_match:
                original_url = url_match.group(1)
            
            # Check for version info (like (2013), (2021), etc.)
            version_match = re.search(r'<br[^>]*>\s*\(([^)]+)\)', cells[0])
            if version_match:
                func_name += ' (' + version_match.group(1) + ')'
        else:
            # Fallback: extract text content from first cell
            func_name = re.sub(r'<[^>]+>', '', cells[0]).strip()
            func_name = re.sub(r'\s+', ' ', func_name)
        
        # Extract description from second cell
        description = cells[1]
        
        # Extract the type (bold text at beginning)
        type_match = re.search(r'<b[^>]*>([^<]+):</b>', description)
        func_type = type_match.group(1) if type_match else ''
        
        # Extract the description text after the type
        desc_text = re.sub(r'<b[^>]*>[^<]+:</b>\s*', '', description)
        desc_text = re.sub(r'<[^>]+>', '', desc_text)  # Remove HTML tags
        desc_text = html.unescape(desc_text)  # Decode HTML entities
        desc_text = re.sub(r'\s+', ' ', desc_text).strip()  # Clean whitespace
        
        # Remove 'This function is not available in Excel for the web.' note
        desc_text = re.sub(r'This function is not available in Excel for the web\.', '', desc_text).strip()
        
        # Create individual function file
        individual_filename = create_individual_function_file(func_name, func_type, desc_text, original_url)
        
        # Escape pipe characters in the content
        func_name_display = func_name.replace('|', '\\|')
        func_type_display = func_type.replace('|', '\\|') if func_type else ''
        desc_text_display = desc_text.replace('|', '\\|')
        
        # Create link to individual function file
        func_name_link = f"[{func_name_display}](functions/{individual_filename})"
        
        function_count += 1
        markdown_lines.append(f'| {function_count} | {func_name_link} | {func_type_display} | {desc_text_display} |')

# Write to markdown file
with open('function_list.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(markdown_lines))

print('Successfully converted HTML table to Markdown!')
print(f'Created {function_count} function entries')
print(f'Created {function_count} individual function files in functions/ directory')
