import re
 from openpyxl import load_workbook
 from openpyxl.utils import coordinate_to_tuple, get_column_letter
 
 def is_within_range(cell, min_cell, max_cell):
     """
     Checks if the cell (e.g., "B2") is within the range defined by [min_cell, max_cell].
     Coordinates are interpreted as (col, row) tuples.
     """
     col, row = coordinate_to_tuple(cell)
     min_col, min_row = coordinate_to_tuple(min_cell)
     max_col, max_row = coordinate_to_tuple(max_cell)
     
     return (min_col <= col <= max_col) and (min_row <= row <= max_row)
 
 def sanitize_function_name(name):
     """
     Sanitize the input name to be used as a valid JavaScript function name.
     Rules:
       - Remove spaces and any invalid characters.
       - Ensure the first character is not a digit (prefix with "_" if needed).
     """
     # Remove any character that is not alphanumeric or underscore or $
     sanitized = re.sub(r'[^A-Za-z0-9_$]', '', name)
     # If empty, use a default name.
     if not sanitized:
         sanitized = "excelFunction"
     # If the first character is a number, prefix with an underscore.
     if re.match(r'^[0-9]', sanitized):
         sanitized = "_" + sanitized
     return sanitized
 
 def excel_sheet_to_js_function(excel_filename, sheet_name=None, min_cell=None, max_cell=None):
     """
     Reads an Excel (.xlsx) file and returns a string containing JavaScript code.
     
     The generated JavaScript function has the same name as the sheet (sanitized to a valid
     JavaScript function identifier) and is defined as:
     
          function FUNCTION_NAME(d) { ... return computed; }
     
     Where d is an object with input cell values.
     
     Any cell with a formula (a string starting with "=") is processed as a computed cell.
     Intermediate computed values are computed in the order they appear in the worksheet
     (using row-order iteration), so that a cell computed earlier can be used later.
     
     Optional limit parameters:
       - min_cell: e.g. "A1" (upper-left cell). Only cells with coordinates not less than min_cell
                   are processed.
       - max_cell: e.g. "D10" (lower-right cell). Only cells with coordinates not greater than max_cell
                   are processed.
     """
     
     # Load the workbook and select the sheet
     wb = load_workbook(excel_filename, data_only=False)
     ws = wb[sheet_name] if sheet_name else wb.active
 
     # Use worksheet title for function name and sanitize it.
     js_func_name = sanitize_function_name(ws.title)
     
     # Regular expression pattern to match Excel cell references (e.g., A1, B12)
     cell_ref_pattern = r"(\b[A-Z]+[0-9]+\b)"
     
     # This list will hold (cell_coordinate, js_expression) for all formula cells.
     computed_cells = []
     
     # Process every cell in the worksheet in row order.
     for row in ws.iter_rows():
         for cell in row:
             cell_coord = cell.coordinate
             # If range limits are provided, skip cells outside the range.
             if min_cell and max_cell:
                 if not is_within_range(cell_coord, min_cell, max_cell):
                     continue
             elif min_cell:
                 # Use Excel's maximum known cell coordinate "XFD1048576" as upper bound.
                 if not is_within_range(cell_coord, min_cell, "XFD1048576"):
                     continue
             elif max_cell:
                 if not is_within_range(cell_coord, "A1", max_cell):
                     continue
                     
             # If the cell contains a formula (data_type "f"), process it.
             if cell.data_type == 'f':
                 formula = cell.value
                 if formula and isinstance(formula, str) and formula.startswith("="):
                     # Remove the leading "=".
                     clean_formula = formula[1:]
                     # Replace Excel cell references (e.g., A1) with JS object's property access: d["A1"]
                     js_expression = re.sub(cell_ref_pattern, r'd["\1"]', clean_formula)
                     computed_cells.append((cell_coord, js_expression))
     
     # Build the JavaScript function code as a list of lines.
     js_lines = [
         f"function {js_func_name}(d) {{",
         "    // d is an object containing input cell values",
         "    let computed = {};"
     ]
     
     # Add computed cell assignments in order. This supports intermediate computations.
     for cell_coord, expression in computed_cells:
         js_lines.append(f"    computed['{cell_coord}'] = {expression};")
     
     js_lines.append("    return computed;")
     js_lines.append("}")
     
     # Return the JS function code as a string.
     js_function_code = "\n".join(js_lines)
     return js_function_code
 
 # ------------------------------
 # Example usage:
 if __name__ == "__main__":
     excel_file = "example.xlsx"  # Replace with the path to your Excel file.
     
     # Optionally limit the processed cells:
     min_cell = "A1"
     max_cell = "D20"
     
     # Specify a sheet name if needed, otherwise the active sheet is used.
     sheet = None  # or e.g., "Sheet1"
     
     js_code = excel_sheet_to_js_function(excel_file, sheet_name=sheet, min_cell=min_cell, max_cell=max_cell)
     print("Generated JavaScript function code:")
     print(js_code)
 
