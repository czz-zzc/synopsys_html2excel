#########################################################################################
# Engineer:         czz
# Date:             2025-10-28
# Version:          1.0
# Description: This script converts new Excel format to old format for gen_reg.py
#########################################################################################

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import argparse
import os
import re

def parse_default_value(default_str, bit_width):
    """
    Convert default value from "0x..." format to Verilog format like "3'h0", "16'd100", "16'h1F"
    
    Args:
        default_str: String like "0x0", "0x64", "0x1F" from new format
        bit_width: Width of the field in bits
    
    Returns:
        Verilog format string like "3'h0", "16'h64", "16'h1F"
    """
    if default_str is None or default_str == "":
        return f"{bit_width}'h0"
    
    default_str = str(default_str).strip()
    
    # If already in Verilog format (width'base value), return as is
    pattern = r"(\d+)'([bdhBDH])([0-9a-fA-F_]+)"
    match = re.match(pattern, default_str)
    if match:
        return default_str
    
    # Parse 0x format and convert to Verilog format
    if default_str.startswith("0x") or default_str.startswith("0X"):
        try:
            int_val = int(default_str, 16)
            # Use hex format for output
            return f"{bit_width}'h{int_val:x}"
        except:
            return f"{bit_width}'h0"
    
    # Try to parse as plain decimal number
    try:
        int_val = int(default_str, 10)
        return f"{bit_width}'h{int_val:x}"
    except:
        pass
    
    # Default fallback
    return f"{bit_width}'h0"

def convert_attribute_to_sw_access(attribute):
    """
    Convert attribute field to sw_access field
    
    Args:
        attribute: String like "RW", "RO", "W1C", "pulse", etc.
    
    Returns:
        Standard sw_access value: "RW", "RO", "W1C", "W1P"
    """
    if attribute is None or attribute == "":
        return "RW"  # Default
    
    attr = str(attribute).strip().upper()
    
    # Map pulse to W1P
    if attr == "PULSE":
        return "W1P"
    elif attr in ["RW", "RO", "W1C", "W1P"]:
        return attr
    else:
        return "RW"  # Default fallback

def calculate_bit_width(bit_str):
    """Calculate bit width from bit range string"""
    if bit_str is None or bit_str == "":
        return 1
    
    bit_str = str(bit_str).strip()
    parts = bit_str.split(':')
    
    if len(parts) == 1:
        return 1
    elif len(parts) == 2:
        start = int(parts[0])
        end = int(parts[1])
        return abs(start - end) + 1
    
    return 1

def parse_new_format(filename):
    """
    Parse new Excel format
    
    Returns:
        module_info: dict with module information
        registers: list of register dictionaries
    """
    wb = openpyxl.load_workbook(filename)
    
    # Assume first sheet has register data, second sheet has variable definitions
    ws_regs = wb.worksheets[0]
    
    module_info = {}
    registers = []
    
    # Parse module info from first few rows (simplified header)
    # Row 1: Module | value
    # Row 2: ID | value  
    # Row 3: Base Addr | value
    for row in ws_regs.iter_rows(max_row=3, values_only=True):
        if row[0] and row[1]:
            key = str(row[0]).strip().lower().replace(' ', '_')
            if key == "module":
                module_info["module"] = row[1]
            elif key == "id":
                # id is addr_width in new format
                module_info["addr_width"] = row[1]
            elif key == "base_addr":
                module_info["base_addr"] = row[1]
    
    # Set defaults if not found
    if "addr_width" not in module_info:
        module_info["addr_width"] = 12
    if "base_addr" not in module_info:
        module_info["base_addr"] = "32'h0000"
    
    # Add missing fields for old format
    module_info["owner"] = "czz"
    module_info["size"] = "4KB"
    module_info["data_width"] = "data_width"
    module_info["cfg_interface"] = "regbus"
    module_info["template_version"] = "v0p0"
    module_info["tool_version"] = "v0p0"
    
    # Parse register definitions
    current_reg = None
    header_found = False
    
    for row_idx, row in enumerate(ws_regs.iter_rows(min_row=4, values_only=True), start=4):
        # Find header row (Row 4: Reg Name | reg_description | offset Addr | Bits | Field name | attribute | field_description | default value)
        if not header_found:
            if row[0] and str(row[0]).strip().lower().replace(' ', '_') in ["reg_name", "reg"]:
                header_found = True
            continue
        
        reg_name = row[0]
        reg_description = row[1] if len(row) > 1 else ""
        offset = row[2] if len(row) > 2 else None
        bits = row[3] if len(row) > 3 else None
        field_name = row[4] if len(row) > 4 else None
        attribute = row[5] if len(row) > 5 else None
        field_description = row[6] if len(row) > 6 else ""
        default_value = row[7] if len(row) > 7 else None
        
        # New register row (has reg_name and offset)
        if reg_name and offset:
            # Save previous register
            if current_reg:
                registers.append(current_reg)
            
            # Convert offset from word address to byte address (multiply by 4)
            # New format uses word (4-byte) addressing, old format uses byte addressing
            offset_str = str(offset).strip()
            
            # Check if it's an expression (e.g., "0x40+i*2" or "0x2a0+m")
            if '+' in offset_str:
                # Parse expression like "0x40+i*2" or "0x2a0+m"
                parts = offset_str.split('+')
                base = parts[0].strip()
                var_part = parts[1].strip()  # "i*2" or "m*2" or "m"
                
                # Convert base to byte addressing
                base_word = int(base, 16)
                base_byte = base_word * 4
                
                # Check if variable part has multiplication
                if '*' in var_part:
                    # Extract variable and step (e.g., "i*2")
                    var_match = re.match(r'([a-zA-Z_]\w*)\*(.+)', var_part)
                    if var_match:
                        variable = var_match.group(1)
                        step_str = var_match.group(2).strip()
                        step_word = int(step_str, 16) if step_str.startswith('0x') else int(step_str, 16)
                        step_byte = step_word * 4
                        offset_byte = f"0x{base_byte:x}+{variable}*0x{step_byte:x}"
                    else:
                        offset_byte = f"0x{base_byte:x}"
                else:
                    # Just variable without multiplication (e.g., "0x2a0+m" means "0x2a0+m*1")
                    variable = var_part
                    offset_byte = f"0x{base_byte:x}+{variable}*0x4"
            else:
                # Simple numeric offset
                word_addr = int(offset_str, 16) if offset_str.startswith('0x') or offset_str.startswith('0X') else int(offset_str, 16)
                byte_addr = word_addr * 4
                offset_byte = f"0x{byte_addr:x}"
            
            # Start new register
            current_reg = {
                "offset": offset_byte,
                "reg_name": str(reg_name).strip(),
                "fields": []
            }
            
            # If this row also has field info, add it
            if field_name:
                bit_width = calculate_bit_width(bits)
                sw_access = convert_attribute_to_sw_access(attribute)
                default_hex = parse_default_value(default_value, bit_width)
                
                current_reg["fields"].append({
                    "bits": str(bits).strip() if bits else "0",
                    "field": str(field_name).strip(),
                    "sw_access": sw_access,
                    "hw_access": "",  # Empty by default
                    "default": default_hex
                })
        
        # Field row (has field_name but no reg_name or offset)
        elif field_name and current_reg:
            bit_width = calculate_bit_width(bits)
            sw_access = convert_attribute_to_sw_access(attribute)
            default_hex = parse_default_value(default_value, bit_width)
            
            current_reg["fields"].append({
                "bits": str(bits).strip() if bits else "0",
                "field": str(field_name).strip(),
                "sw_access": sw_access,
                "hw_access": "",  # Empty by default
                "default": default_hex
            })
    
    # Save last register
    if current_reg:
        registers.append(current_reg)
    
    # Check if there's a second sheet with variable definitions
    var_ranges = {}
    if len(wb.worksheets) > 1:
        ws_vars = wb.worksheets[1]
        for row in ws_vars.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                break
            name = row[0]
            range_str = row[1]
            if range_str:
                range_str = str(range_str)
                if '~' in range_str:
                    _, max_val = map(int, range_str.split('~'))
                    var_ranges[name] = max_val
                elif '-' in range_str:
                    _, max_val = map(int, range_str.split('-'))
                    var_ranges[name] = max_val
    
    return module_info, registers, var_ranges

def write_old_format(module_info, registers, var_ranges, output_file):
    """
    Write data to old Excel format
    """
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    
    # Create main register sheet
    ws_regs = wb.create_sheet("registers", 0)
    
    # Header style
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    
    # Write module information (rows 1-9)
    module_data = [
        ["module", module_info.get("module", "")],
        ["owner", module_info.get("owner", "czz")],
        ["size", module_info.get("size", "4KB")],
        ["cfg_interface", module_info.get("cfg_interface", "regbus")],
        ["base_addr", module_info.get("base_addr", "32'h0000")],
        ["addr_width", module_info.get("addr_width", 12)],
        ["data_width", module_info.get("data_width", "data_width")],
        ["template_version", module_info.get("template_version", "v0p0")],
        ["tool_version", module_info.get("tool_version", "v0p0")]
    ]
    
    for row_idx, (key, value) in enumerate(module_data, start=1):
        ws_regs.cell(row=row_idx, column=1, value=key)
        ws_regs.cell(row=row_idx, column=2, value=value)  # Column B
    
    # Write register table header (row 10)
    headers = ["offset", "reg_name", "bits", "field", "sw_access", "hw_access", "default", "attibute", "description"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws_regs.cell(row=10, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    
    # Write register data (starting from row 11)
    current_row = 11
    for reg in registers:
        offset = reg["offset"]
        reg_name = reg["reg_name"]
        
        for field_idx, field in enumerate(reg["fields"]):
            # First field row includes offset and reg_name
            if field_idx == 0:
                ws_regs.cell(row=current_row, column=1, value=offset)
                ws_regs.cell(row=current_row, column=2, value=reg_name)
            
            # Field data
            ws_regs.cell(row=current_row, column=3, value=field["bits"])
            ws_regs.cell(row=current_row, column=4, value=field["field"])
            ws_regs.cell(row=current_row, column=5, value=field["sw_access"])
            ws_regs.cell(row=current_row, column=6, value=field["hw_access"])
            ws_regs.cell(row=current_row, column=7, value=field["default"])
            
            current_row += 1
    
    # Create variable definitions sheet if needed
    if var_ranges:
        ws_vars = wb.create_sheet("variables", 1)
        
        # Header
        ws_vars.cell(row=1, column=1, value="name").font = header_font
        ws_vars.cell(row=1, column=2, value="range").font = header_font
        
        # Data
        for idx, (name, max_val) in enumerate(var_ranges.items(), start=2):
            ws_vars.cell(row=idx, column=1, value=name)
            ws_vars.cell(row=idx, column=2, value=f"0~{max_val}")
    
    # Adjust column widths
    ws_regs.column_dimensions['A'].width = 15
    ws_regs.column_dimensions['B'].width = 20
    ws_regs.column_dimensions['C'].width = 10
    ws_regs.column_dimensions['D'].width = 25
    ws_regs.column_dimensions['E'].width = 12
    ws_regs.column_dimensions['F'].width = 12
    ws_regs.column_dimensions['G'].width = 12
    
    # Save workbook
    wb.save(output_file)
    print(f"Successfully converted to old format: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert new Excel format to old format')
    parser.add_argument('input', help='Input Excel file (new format)')
    parser.add_argument('-o', '--output', help='Output Excel file (old format). Default: <input>_old.xlsx')
    args = parser.parse_args()
    
    # Determine output filename
    if args.output:
        output_file = args.output
        # If output is a directory, generate filename automatically
        if os.path.isdir(output_file):
            input_basename = os.path.basename(args.input)
            base_name = os.path.splitext(input_basename)[0]
            output_file = os.path.join(output_file, f"{base_name}_old.xlsx")
    else:
        base_name = os.path.splitext(args.input)[0]
        output_file = f"{base_name}_old.xlsx"
    
    # Convert
    module_info, registers, var_ranges = parse_new_format(args.input)
    write_old_format(module_info, registers, var_ranges, output_file)
