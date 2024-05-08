# Excel Formula Extractor 📊

This tool is designed to extract formulas, constants, and exceptions from an Excel workbook. It provides insights into the calculations made within the workbook, aiding in analysis and understanding.

## Installation
1. Clone or download the repository.
2. Ensure you have Python 3.x installed.
3. Install the required dependencies:
    ```
    pip install openpyxl
    ```
4. Run the script.

## How to Use
1. Execute the script.
2. Select the Excel file you want to analyze.
3. Choose whether to run a full analysis or specify a starting cell.
    - For a full analysis, the tool will extract data from predefined starting cells.
    - For a specific analysis, input the sheet name and cell location to start the extraction.

## Running a Full Analysis
- The tool iterates through predefined starting cells, extracting formulas, constants, and exceptions.
- Results are saved to separate files named according to the data type and cell position.

## Specific Analysis
- Specify the sheet name and cell location to begin the analysis.
- Results are displayed for the provided cell and its dependencies.

## Output
- The tool generates output files containing extracted formulas, constants, and exceptions.
- Output files are saved in the same directory as the script.

## Example Usage
```python
python excel_formula_extractor.py
