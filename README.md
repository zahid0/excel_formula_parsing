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

### Function Parameters

The `excel_sheet_to_js_function` takes the following parameters:

- `excel_filename`: Path to your Excel file (.xlsx)
- `sheet_name`: Name of the sheet to process (optional, default is the active sheet)
- `min_cell`: Minimum cell coordinate to process (optional)
- `max_cell`: Maximum cell coordinate to process (optional)

### Function Returns

The function returns:
1. A string containing the JavaScript function code
2. The name of the generated JavaScript function
3. A string containing the input object `d` with input cell values

### Example Usage

```python
excel_file = "test_formula_sheet.xlsx"
sheet_name = "Sheet1"  # or other sheet name

js_function_code, function_name, input_object = excel_sheet_to_js_function(
    excel_file,
    sheet_name=sheet_name,
    min_cell="A1",
    max_cell="D20"
)

print("// JavaScript Function:")
print(js_function_code)
print("\n// Input Object:")
print(input_object)
print("\n// Example usage:")
print(f"console.log({function_name}(d))")
```

### Output

The generated JavaScript function will look like this:

```javascript
functionSheet1(d) {
    // d is an object containing input cell values
    let computed = {};
    computed['B2'] = d['A1'] * 2;
    computed['C3'] = computed['B2'] + d['B1'];
    return computed;
}
```

The input object will look like this:

```javascript
const d = {
    "A1": 10,
    "B1": 5,
};
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
