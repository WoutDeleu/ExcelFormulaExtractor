# Excel Formula Extractor 📊

This tool is designed to extract formulas, constants, and exceptions from an Excel workbook, in this case specifically designed for the file _PB-berekeningen.xlsx_. It provides insights into the calculations made within the workbook, aiding in analysis and understanding. It extracts formulas, and writes them in a more readable, standard syntax, all combined in one file.

The idea is to summarize and analyse complex and long excel files withouth having to click through all the involved cells manually!

## Installation 
1. Clone or download this repository, and enter the folder.
2. Ensure you have Python 3.x installed and Python added to the path of your device, as well as pip configured.
3. Install the required dependencies using the following command:
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
2. Select the Excel file you want to analyze (this can take some time with large ). 
3. Choose whether to run a full analysis or specify a starting cell.
    - For a full analysis, the tool will extract data from predefined starting cells.
    - For a specific analysis, input the sheet name and cell location to start the extraction as an answer to the prompt.

### Command Line arguments
```
--full_analysis
    Resolves all the formulas, errors and variables from a predefined list of cells!

--single_cell
    Resolves all the formulas, errors and variables for one chosen cell!

--file, -f [path to file]
    Pass on the path of the file you want to analyse in advance (so withouth having to wait for the prompt).

--cell, -c [cellnumber]
    The excel cell you want to start from in case of the default configuration (default configuration = NO full analysis). E.g. 'C34'.

--sheetname, -sh [sheetname]
    The sheetname of the excel cell you want to analyse

--write_to_file, -wtf
    A flag indicating if the results should be written to the corresponding files! This 
```
*_REMARK_*: 
you can't provide just the --cell argument or just the --sheetname argument. Either both, or none, otherwise the program will fail. Same goes for the --full_analysis flag and the --single_cell flag!

## How it works
which EXCEL functions:..


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
