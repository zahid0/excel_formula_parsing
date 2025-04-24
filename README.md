# excel-to-js-function

A Python tool to convert Excel formulas into JavaScript functions.

## Overview

This repository provides a Python function `excel_sheet_to_js_function` that converts Excel spreadsheet formulas into equivalent JavaScript functions. It parses an Excel file, extracts formulas and their dependencies, and generates a JavaScript function that maintains the same computation logic.

The tool creates:
1. A JavaScript function that calculates values based on Excel formulas
2. An input object containing only the necessary input cells used in the calculations

The generated JavaScript function can be used in web applications to replicate the computation logic originally implemented in Excel.

## Setup

### Prerequisites
- Python 3.8+
- pip (Python package installer)

### Linux Setup

1. Install Python and pip if not already installed:
   ```bash
   sudo apt-get install python3 python3-pip
   ```

2. Install virtualenv (optional but recommended):
   ```bash
   sudo pip3 install virtualenv
   ```

3. Create a virtual environment and activate it (optional):
   ```bash
   virtualenv myenv
   source myenv/bin/activate
   ```

### Mac Setup

1. Install Python and pip if not already installed using Homebrew:
   ```bash
   brew install python
   ```

2. Install virtualenv (optional but recommended):
   ```bash
   pip install virtualenv
   ```

3. Create a virtual environment and activate it (optional):
   ```bash
   virtualenv myenv
   source myenv/bin/activate
   ```

### Windows Setup

1. Install Python 3.x from the official Python website.
   - Make sure to add Python to your PATH during installation.

2. Install virtualenv using pip:
   ```cmd
   pip install virtualenv
   ```

3. Create a virtual environment and activate it (optional):
   ```cmd
   virtualenv myenv
   myenv\Scripts\activate
   ```

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-to-js-function.git
   cd excel-to-js-function
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

This tool converts Excel sheet formulas into JavaScript functions that can be used in web applications. It processes specified cells or ranges of cells and generates corresponding JavaScript code.

## Command Line Arguments

The tool accepts the following command-line arguments:

### Mandatory Argument
- `excel_file`: Path to your Excel file (`.xlsx`)

### Optional Arguments
- `--sheet`: Comma-separated list of sheet names to process
  - If not provided, all sheets in the Excel file will be processed
  - Example: `--sheet "Sheet1,Sheet2"`

- `--min_cell`: Minimum cell coordinate to limit processing
  - Example: `--min_cell A1`

- `--max_cell`: Maximum cell coordinate to limit processing
  - Example: `--max_cell D20`

- `--include-test-code`: Include test code for verification
  - When present, includes d_object and console.log() statements
  - By default, only the JavaScript function code is printed
  - Example: `--include-test-code`

### Example Usage
```bash
python excel_to_js.py example.xlsx --sheet "Sheet1,Sheet2" --min_cell A1 --max_cell D20
```

Or process all sheets without cell limits:
```bash
python excel_to_js.py example.xlsx
```

To include test code for verification:
```bash
python excel_to_js.py example.xlsx --include-test-code
```

## Explanation

### Sheet Processing
- If you specify `--sheet "Sheet1,Sheet2"`, only those sheets will be processed
- If you omit `--sheet`, all sheets in the Excel file are processed

### Cell Range Limitation
- Use `--min_cell` and `--max_cell` to limit the range of cells processed
  - Example: `--min_cell A1 --max_cell D20` will process cells from A1 to D20

### Test Code Inclusion
- By default, only the JavaScript function code is printed
- When `--include-test-code` is used, additional code including:
  - A JavaScript object containing the formula data
  - Example usage with console.log() statements
- is included in the output

## Common Use Cases

1. Convert formulas from a specific sheet:
```bash
python excel_to_js.py example.xlsx --sheet "Sheet1"
```

2. Convert formulas from multiple sheets:
```bash
python excel_to_js.py example.xlsx --sheet "Sheet1,Sheet2,Sheet3"
```

3. Convert formulas within specific cell range:
```bash
python excel_to_js.py example.xlsx --min_cell A1 --max_cell D20
```

4. Convert all sheets without cell range limits:
```bash
python excel_to_js.py example.xlsx
```

5. Include test code for verification:
```bash
python excel_to_js.py example.xlsx --include-test-code
```

## Output

The tool generates JavaScript code for each processed sheet. By default, it will only print the JavaScript function code. When `--include-test-code` is used, it also prints additional code for testing purposes.

Default output (without `--include-test-code`):
```javascript
// Code for sheet: Sheet1
function ConvertExcelSheet(d) {
    // ... [JavaScript code]
}
```

Output with `--include-test-code`:
```javascript
// Code for sheet: Sheet1
function ConvertExcelSheet(d) {
    // ... [JavaScript code]
}
const d_object = { /* ... data object ... */ };
console.log(ConvertExcelSheet(d));
```

## How It Works

1. **Cell Parsing**: The function parses the Excel sheet and identifies all cells that contain formulas.
2. **Dependency Collection**: For each formula, it collects all referenced cells and builds a dependency graph.
3. **Topological Sort**: It performs a topological sort on the dependency graph to ensure proper computation order.
4. **JavaScript Code Generation**: It converts the Excel formulas into JavaScript expressions, substituting cell references with either computed values or input values.
5. **Input Object Generation**: It creates an input object containing only the cells that need to be provided as input to the JavaScript function.

## Limitations

- Currently supports cell reference formulas only
- Does not support all Excel functions (only supports cell references and basic operations)
- Cannot handle circular dependencies in formulas
