# Excel Formula Extractor 📊

This tool is designed to extract formulas, constants, and exceptions from an Excel workbook, in this case specifically designed for the file _PB-berekeningen.xlsx_. It provides insights into the calculations made within the workbook, aiding in analysis and understanding. It extracts formulas, and writes them in a more readable, standard syntax, all combined in one file.

The idea is to summarize and analyse complex and long excel files withouth having to click through all the involved cells manually!

## Installation 
1. Clone or download the repository.
2. Ensure you have Python 3.x installed and Python added to the path of your device.
3. Install the required dependencies:
    ```shell
    pip install -r requirements.txt
    ```
4. Run the script with the desired flags

## How to Use
### Without command line arguments
1. Execute the script without flag options.
    ```shell
    python main.py
    ```
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

## Contribution