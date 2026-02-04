#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import re
import os
import argparse
import glob
import openpyxl
from openpyxl.styles import Alignment

def format_description(td_tag):
    """
    Format the description cell to preserve lists and paragraphs
    while cleaning up whitespace.
    """
    lines = []
    
    def process_element(element):
        if isinstance(element, str): # NavigableString
            text = element.strip()
            if text:
                lines.append(text)
            return

        if element.name == 'br':
            lines.append('\n')
            return

        if element.name in ['ul', 'ol']:
            lines.append('\n')
            for li in element.find_all('li'):
                # Clean up list item text
                li_text = ' '.join(li.get_text(separator=' ', strip=True).split())
                lines.append(f"• {li_text}")
                lines.append('\n')
            lines.append('\n')
            return

        if element.name in ['p', 'div']:
            # If the block contains a list, iterate through children to preserve structure
            if element.find(['ul', 'ol']):
                lines.append('\n')
                for child in element.children:
                    process_element(child)
                lines.append('\n')
            else:
                # Regular text block
                p_text = ' '.join(element.get_text(separator=' ', strip=True).split())
                if p_text:
                    lines.append(p_text)
                    lines.append('\n\n')
            return

        # Handle other wrapper tags (span, strong, etc.)
        text = ' '.join(element.get_text(separator=' ', strip=True).split())
        if text:
            lines.append(text)

    # Main loop processing top-level children
    for child in td_tag.children:
        process_element(child)

    # Join and clean up
    # Join with space usually, but respect \n logic
    result = ""
    for line in lines:
        if line == '\n' or line == '\n\n':
            result += line
        else:
            if result and not result.endswith('\n') and not result.endswith(' '):
                 result += " "
            result += line
            
    # Final cleanup of excessive newlines
    import re
    return re.sub(r'\n{3,}', '\n\n', result).strip()

def parse_html_file(html_file, excel_data):
    # Read HTML file content
    with open(html_file, 'r', encoding='ISO-8859-1') as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, 'html.parser')
    # Output HTML structure to txt for debugging
    #debug_txt = os.path.splitext(html_file)[0] + '_soup_output.txt'
    #with open(debug_txt, 'w', encoding='utf-8') as f:
    #    f.write(soup.prettify())

    current_register = None

    # Find all register definitions
    for element in soup.find_all(['p', 'div']):
        # Detect register title
        if element.name == 'p' and element.find('strong') and element.find('a'):
            #if strong and any string in strong.text
            #if strong and re.search(r'\b[A-Z_]+\b', strong.text, re.IGNORECASE):
            strong = element.find('strong')
            # Start a new register
            current_register = {
                'name': strong.get_text(strip=True),
                'offset': None,
                'size': None,
                'description': None,
                'fields': []
            }
            # Find basic info list
            next_ul = element.find_next('ul', class_='itemizedlist')
            if next_ul:
                for li in next_ul.find_all('li', class_='listitem'):
                    text = li.get_text(strip=True)
                    if 'Offset:' in text:
                        current_register['offset'] = text.split(':')[-1].strip()
                    if 'Size:' in text:
                        current_register['size'] = text.split(':')[-1].strip()
                    if 'Description:' in text:
                        current_register['description'] = text.split(':', 1)[-1].strip()

        # Detect field table
        elif current_register and element.name == 'div' and 'table' in element.get('class', []):
            table = element.find('table', class_='table')
            if not table:
                continue

            # Process table rows
            for row in table.find_all('tr')[1:]:  # Skip header
                cols = row.find_all('td')
                if len(cols) < 4:
                    continue
                default_value = None
                # Directly find tag containing "Value After Reset"
                value_tag = cols[3].find(lambda tag: tag.name == 'p' and 
                                         'Value After Reset' in tag.get_text())
                if value_tag:
                    value_text = value_tag.get_text(strip=True)
                    if 'Value After Reset:' in value_text:
                        default_value = value_text.split('Value After Reset:', 1)[-1].split('\n')[0].strip()
                    
                # Extract field info
                field = {
                    'bits': cols[0].get_text(strip=True),
                    'name': cols[1].get_text(strip=True),
                    'access': cols[2].get_text(strip=True),
                    'description': format_description(cols[3]),
                    'default': default_value
                }
                if not field['name'] or "reserved" in field['name'].lower():
                    continue
                current_register['fields'].append(field)
                # Add to Excel data structure
                # If register name first time, print offset and reg_name
                if len(current_register['fields']) == 1:
                    print(f"Offset: {current_register['offset']},Register: {current_register['name']}")
                # Append data to excel_data
                excel_data['offset'].append(current_register['offset'] if len(current_register['fields']) == 1 else "")
                excel_data['reg_name'].append(current_register['name'] if len(current_register['fields']) == 1 else "")
                # Append field data
                excel_data['bits'].append(field['bits'])
                if 'reserved' in field['name'].lower():
                    excel_data['field'].append("Reserved")
                else:
                    excel_data['field'].append(field['name'])

                if 'R/W' in field['access']:
                    excel_data['sw_access'].append("RW")
                elif 'R' in field['access']:
                    excel_data['sw_access'].append("RO")
                else:
                    excel_data['sw_access'].append("unknown")

                excel_data['hw_access'].append("")  # This column is empty in the example
                # Format default value: convert 0xaa to 8'ha format, where 8 is from field['bits']
                def format_default(bits, default):
                    if default and re.match(r"^0x[0-9a-fA-F]+$", default):
                        # Extract width from bits, e.g., "7:0" -> 8, "3" -> 1
                        m = re.match(r"(\d+):(\d+)", bits)
                        if m:
                            width = int(m.group(1)) - int(m.group(2)) + 1
                        elif bits.isdigit():
                            width = 1
                        else:
                            width = ""
                        return f"{width}'h{default[2:]}" if width else default
                    return default or ""

                excel_data['default'].append(format_default(field['bits'], field['default']))  # Avoid None values
                excel_data['attribute'].append("normal")  # All are "normal" in the example
                excel_data['description'].append(field['description'])  # Populate description

def write_excel_with_header(excel_data, output_file):
    # Get module name from output file name (remove extension)
    module_name = os.path.splitext(os.path.basename(output_file))[0]

    # Custom header info (first 9 rows)
    custom_header = [
        ['module', module_name],
        ['owner', 'czz'],
        ['size', '64KB'],
        ['cfg_interface', 'regbus'],
        ['base_addr', "32'h0000"],
        ['addr_width', '16'],
        ['data_width', '32'],
        ['template_version', 'v0p0'],
        ['tool_version', 'v0p0'],
    ]
    # Table header (row 10)
    table_header = ['offset', 'reg_name', 'bits', 'field', 'sw_access', 'hw_access', 'default', 'attribute', 'description']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'reglist'

    # Write custom header and merge columns B-I for each row
    for row_idx, row in enumerate(custom_header, 1):
        ws.cell(row=row_idx, column=1, value=row[0])
        ws.cell(row=row_idx, column=2, value=row[1])
        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=9)

    # Write table header
    for col_idx, value in enumerate(table_header, 1):
        ws.cell(row=10, column=col_idx, value=value)

    # Write data starting from row 11
    # Transpose data from columns to rows
    rows = zip(
        excel_data['offset'], 
        excel_data['reg_name'], 
        excel_data['bits'], 
        excel_data['field'], 
        excel_data['sw_access'], 
        excel_data['hw_access'], 
        excel_data['default'], 
        excel_data['attribute'], 
        excel_data['description']
    )

    for r_idx, row in enumerate(rows, 11):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if c_idx == 9: # Description column
                cell.alignment = Alignment(wrap_text=True, vertical='top')
            else:
                cell.alignment = Alignment(vertical='top')

    # Auto-adjust column width and row height
    description_width = 80
    
    for col in ws.columns:
        max_length = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        
        if col_letter == 'I': # Description column fixed width
            ws.column_dimensions[col_letter].width = description_width
            continue

        for cell in col:
            try:
                if cell.value:
                    # Limit max width for other columns
                    cell_len = len(str(cell.value))
                    if cell_len > 50: cell_len = 50 
                    max_length = max(max_length, cell_len)
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Adjust row height for description
    for row in ws.iter_rows(min_row=11):
        max_lines = 1
        cell = row[8] # Description is 9th column (index 8)
        if cell.value:
            text = str(cell.value)
            # Estimate lines: count newlines + wrapped lines
            lines = 0
            for paragraph in text.split('\n'):
                # Ceiling division for wrapping
                para_len = len(paragraph)
                if para_len == 0:
                    lines += 1 # Empty line
                else:
                    lines += (para_len + description_width - 1) // description_width
            max_lines = max(max_lines, lines)
        
        # Approximate height: 15 points per line
        if max_lines > 1:
            ws.row_dimensions[row[0].row].height = max_lines * 15
        
    wb.save(output_file)

def main():
    parser = argparse.ArgumentParser(description='Extract register info from one or more HTML files and export to Excel.')
    parser.add_argument('html_files', nargs='+', help='HTML file(s) to parse. Wildcards supported (e.g., *.html)')
    parser.add_argument('-o', '--output', default='register_output.xlsx', help='Output Excel file name')
    args = parser.parse_args()

    # Prepare data structure
    excel_data = {
        'offset': [],
        'reg_name': [],
        'bits': [],
        'field': [],
        'sw_access': [],
        'hw_access': [],
        'default': [],
        'attribute': [],
        'description': []
    }

    # Support wildcards for input files
    html_file_list = []
    for pattern in args.html_files:
        html_file_list.extend(glob.glob(pattern))

    if not html_file_list:
        print("No HTML files found.")
        return

    for html_file in html_file_list:
        print(f"Parsing {html_file} ...")
        parse_html_file(html_file, excel_data)

    # Write with custom header and merged cells
    write_excel_with_header(excel_data, args.output)
    print(f"Excel file saved to {args.output}")

if __name__ == '__main__':
    main()
