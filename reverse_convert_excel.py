#########################################################################################
# Engineer:         czz
# Date:             2025-12-16
# Version:          1.0
# Description: This script converts old Excel format (gen_reg.py) back to new Excel format
#########################################################################################

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import argparse
import os
import re

def convert_verilog_default_to_hex(default_str):
    """
    Convert default value from Verilog format like "3'h0", "16'd100", "16'h1F" to "0x..." format
    
    Args:
        default_str: String like "3'h0", "16'h64", "16'h1F"
    
    Returns:
        Hex format string like "0x0", "0x64", "0x1F"
    """
    if default_str is None or default_str == "":
        return "0x0"
    
    default_str = str(default_str).strip()
    
    # Check for Verilog format (width'base value)
    # e.g. 32'h1A, 1'b1, 10'd100
    pattern = r"(\d+)'([bdhBDH])([0-9a-fA-F_]+)"
    match = re.match(pattern, default_str)
    
    if match:
        base = match.group(2).lower()
        value_str = match.group(3).replace('_', '')
        
        try:
            if base == 'h':
                int_val = int(value_str, 16)
            elif base == 'd':
                int_val = int(value_str, 10)
            elif base == 'b':
                int_val = int(value_str, 2)
            else:
                return default_str
                
            return f"0x{int_val:x}"
        except:
            return default_str
            
    # If it's already a number or hex
    if default_str.startswith("0x") or default_str.startswith("0X"):
        return default_str.lower()
        
    if default_str.isdigit():
        return f"0x{int(default_str):x}"
        
    return default_str

def convert_sw_access_to_attribute(sw_access):
    """
    Convert sw_access field to attribute field
    
    Args:
        sw_access: String like "RW", "RO", "W1C", "W1P"
    
    Returns:
        Attribute string: "RW", "RO", "W1C", "PULSE"
    """
    if sw_access is None or sw_access == "":
        return "RW"
    
    access = str(sw_access).strip().upper()
    
    if access == "W1P":
        return "PULSE"
    else:
        return access

def convert_offset_to_word_addr(offset_str, keep_byte_addr=False):
    """
    Convert byte address offset to word address offset (divide by 4)
    If keep_byte_addr is True, keep as byte address.
    Handles expressions like "0x100", "0x100+i*0x10", "(0x0080*i)+0x1100"
    """
    if offset_str is None:
        return "0x0"
        
    offset_str = str(offset_str).strip()

    def replace_num(match):
        num_str = match.group(0)
        try:
            # Handle hex
            if num_str.lower().startswith('0x'):
                val = int(num_str, 16)
            # Handle decimal
            else:
                val = int(num_str, 10)
            
            # Apply conversion
            new_val = val if keep_byte_addr else val // 4
            return f"0x{new_val:x}"
        except:
            return num_str

    # Regex to capture Hex numbers OR Decimal numbers
    # Hex: 0x[0-9a-fA-F]+
    # Decimal: \d+ (ensure it's not part of a variable name like ch1)
    # Using negative lookbehind/lookahead for decimal to avoid matching inside identifiers
    # Matches 0x123 OR 123 (but not a123 or 123a)
    pattern = r'0x[0-9a-fA-F]+|(?<![a-zA-Z_])\d+(?![a-zA-Z_])'
    
    new_offset = re.sub(pattern, replace_num, offset_str, flags=re.IGNORECASE)
    
    # Optional: cleanup *0x1 or 0x1* to simplify i*1 to i
    if not keep_byte_addr:
        # match *0x1 followed by end of string or operator or space (not followed by hex digit)
        new_offset = re.sub(r'\*0x1(?![0-9a-fA-F])', '', new_offset)
        new_offset = re.sub(r'(?<![0-9a-fA-F])0x1\*', '', new_offset)

    return new_offset

def parse_old_format(filename):
    """
    Parse old Excel format
    """
    wb = openpyxl.load_workbook(filename)
    
    # 1. Parse Registers Sheet
    if "registers" in wb.sheetnames:
        ws_regs = wb["registers"]
    else:
        ws_regs = wb.worksheets[0] # Fallback
        
    module_info = {}
    registers = []
    
    # Parse module info (Rows 1-9)
    for row in ws_regs.iter_rows(min_row=1, max_row=9, values_only=True):
        if row[0] and row[1]:
            key = str(row[0]).strip()
            value = row[1]
            module_info[key] = value
            
    # Parse registers (Row 11+)
    # Header is at Row 10: offset, reg_name, bits, field, sw_access, hw_access, default, attibute, description
    
    current_reg_offset = None
    current_reg_name = None
    current_reg_fields = []
    
    # We need to group fields into registers based on offset/reg_name
    # In old format, offset and reg_name are only present in the first row of the register (usually)
    # But convert_excel.py writes them only for the first field.
    
    for row in ws_regs.iter_rows(min_row=11, values_only=True):
        # Stop if empty row (assuming end of table)
        if not any(row):
            continue
            
        offset = row[0]
        reg_name = row[1]
        bits = row[2]
        field = row[3]
        sw_access = row[4]
        # hw_access = row[5] # Not used in new format
        default = row[6]
        # attribute = row[7] # Not used/populated usually
        description = row[8] if len(row) > 8 else ""
        
        # If offset/reg_name is present, it's a new register (or explicit continuation)
        if offset is not None and reg_name is not None:
            # Save previous register
            if current_reg_name is not None:
                registers.append({
                    "offset": current_reg_offset,
                    "reg_name": current_reg_name,
                    "fields": current_reg_fields
                })
            
            current_reg_offset = offset
            current_reg_name = reg_name
            current_reg_fields = []
            
        # Add field to current register
        if field:
            current_reg_fields.append({
                "bits": bits,
                "field": field,
                "sw_access": sw_access,
                "default": default,
                "description": description
            })
            
    # Save last register
    if current_reg_name is not None:
        registers.append({
            "offset": current_reg_offset,
            "reg_name": current_reg_name,
            "fields": current_reg_fields
        })
        
    # 2. Parse Variables Sheet
    var_ranges = {}
    if "variables" in wb.sheetnames:
        ws_vars = wb["variables"]
        for row in ws_vars.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                name = row[0]
                range_str = row[1]
                # Parse "0~10" back to max value
                if '~' in str(range_str):
                    parts = str(range_str).split('~')
                    var_ranges[name] = parts[1]
                elif '-' in str(range_str):
                    parts = str(range_str).split('-')
                    var_ranges[name] = parts[1]
                else:
                    var_ranges[name] = range_str
                    
    return module_info, registers, var_ranges

def write_new_format(module_info, registers, var_ranges, output_file, keep_byte_addr=False):
    """
    Write data to new Excel format
    """
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
        
    # Create Register Sheet
    ws_regs = wb.create_sheet("Sheet1", 0)
    
    # Styles
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    
    # 1. Write Module Info (Rows 1-3)
    # Row 1: Module | value
    ws_regs.cell(row=1, column=1, value="Module")
    ws_regs.cell(row=1, column=2, value=module_info.get("module", ""))
    
    # Row 2: ID | value (addr_width)
    ws_regs.cell(row=2, column=1, value="ID")
    ws_regs.cell(row=2, column=2, value=module_info.get("addr_width", 12))
    
    # Row 3: Base Addr | value
    ws_regs.cell(row=3, column=1, value="Base Addr")
    ws_regs.cell(row=3, column=2, value=module_info.get("base_addr", "32'h0000"))
    
    # 2. Write Header (Row 4)
    headers = ["Reg Name", "Reg_Description", "Offset_Addr", "Bits", "Field Name", "Attribute", "Field_Description", "Default Value"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws_regs.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        
    # 3. Write Registers
    current_row = 5
    for reg in registers:
        # Convert offset
        offset_byte = reg["offset"]
        offset_word = convert_offset_to_word_addr(offset_byte, keep_byte_addr)
        reg_name = reg["reg_name"]
        
        # We don't have reg_description in old format explicitly separated, 
        # but we might have description in the first field which could be reg description?
        # For now, leave reg_description empty or use first field's description if it looks like one?
        # Let's leave it empty as per convert_excel.py behavior (it didn't save it).
        reg_desc = "" 
        
        # Write Register Row (Reg Name, Desc, Offset) - Standalone row
        ws_regs.cell(row=current_row, column=1, value=reg_name)
        ws_regs.cell(row=current_row, column=2, value=reg_desc)
        ws_regs.cell(row=current_row, column=3, value=offset_word)
        current_row += 1

        for field in reg["fields"]:
            # Field data
            ws_regs.cell(row=current_row, column=4, value=field["bits"])
            ws_regs.cell(row=current_row, column=5, value=field["field"])
            
            # Attribute
            attr = convert_sw_access_to_attribute(field["sw_access"])
            ws_regs.cell(row=current_row, column=6, value=attr)
            
            # Description
            cell_desc = ws_regs.cell(row=current_row, column=7, value=field["description"])
            cell_desc.alignment = Alignment(wrap_text=True, vertical='top')
            
            # Default Value
            def_val = convert_verilog_default_to_hex(field["default"])
            cell_def = ws_regs.cell(row=current_row, column=8, value=def_val)
            cell_def.alignment = Alignment(vertical='top')

            # Set alignment for other cells in row
            for col in range(1, 7):
                ws_regs.cell(row=current_row, column=col).alignment = Alignment(vertical='top')
            
            current_row += 1
            
    # 4. Write Variables Sheet (if needed)
    if var_ranges:
        ws_vars = wb.create_sheet("Sheet2", 1)
        
        # Header
        ws_vars.cell(row=1, column=1, value="name")
        ws_vars.cell(row=1, column=2, value="range")
        
        # Data
        for idx, (name, max_val) in enumerate(var_ranges.items(), start=2):
            ws_vars.cell(row=idx, column=1, value=name)
            ws_vars.cell(row=idx, column=2, value=f"0~{max_val}")
            
    # Adjust column widths
    ws_regs.column_dimensions['A'].width = 20
    ws_regs.column_dimensions['B'].width = 20
    ws_regs.column_dimensions['C'].width = 15
    ws_regs.column_dimensions['D'].width = 10
    ws_regs.column_dimensions['E'].width = 20
    ws_regs.column_dimensions['F'].width = 10
    ws_regs.column_dimensions['G'].width = 80
    ws_regs.column_dimensions['H'].width = 15
    
    wb.save(output_file)
    print(f"Successfully converted to new format: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert old Excel format back to new Excel format')
    parser.add_argument('input', help='Input Excel file (old format)')
    parser.add_argument('-o', '--output', help='Output Excel file (new format). Default: <input>_new.xlsx')
    parser.add_argument('--keep-byte-addr', action='store_true', help='Keep address as byte address (do not divide by 4)')
    args = parser.parse_args()
    
    # Determine output filename
    if args.output:
        output_file = args.output
        if os.path.isdir(output_file):
            input_basename = os.path.basename(args.input)
            base_name = os.path.splitext(input_basename)[0]
            # Remove _old suffix if present
            if base_name.endswith('_old'):
                base_name = base_name[:-4]
            output_file = os.path.join(output_file, f"{base_name}_new.xlsx")
    else:
        base_name = os.path.splitext(args.input)[0]
        # Remove _old suffix if present
        if base_name.endswith('_old'):
            base_name = base_name[:-4]
        output_file = f"{base_name}_new.xlsx"
    
    # Convert
    module_info, registers, var_ranges = parse_old_format(args.input)
    write_new_format(module_info, registers, var_ranges, output_file, args.keep_byte_addr)
