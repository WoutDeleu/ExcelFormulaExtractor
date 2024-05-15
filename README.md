# Excel Formula Extractor 📊 📈 📉

This tool is designed to extract formulas, constants, and exceptions from an Excel workbook, in this case specifically designed for the file _PB-berekeningen.xlsx_. It provides insights into the calculations made within the workbook, aiding in analysis and understanding. It extracts formulas, and writes them in a more readable, standard syntax, all combined in one file.

The idea is to summarize and analyse complex and long excel files withouth having to click through all the involved cells manually!

### Contents
- [Installation](#Installation)
- [How To Use](#How-to-Use)
    - [Without command line arguments](#Without-command-line-arguments)
    - [Command line arguments](#Command-line-arguments)
- [How it works](#How-it-works)
- [Output](#Output)
- [Contribution](#Contribution)


## Installation 🐍
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

### Command line arguments
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
    The sheetname of the excel cell you want to analyse.

--write_to_file, -wtf
    A flag indicating if the results should be written to the corresponding files! This will create 3 files per cell (..._exceptions, ..._formulas, ..._values).
```
*_REMARK_*: 
you can't provide just the --cell argument or just the --sheetname argument. Either both, or none, otherwise the program will fail. Same goes for the --full_analysis flag and the --single_cell flag!

## How it works 🧑‍🏫
which EXCEL functions:..


### Running a Full Analysis
- The tool iterates through predefined starting cells, extracting formulas, constants, and exceptions.
- Results are saved to separate files named according to the data type and cell position.

### Specific Analysis
- Specify the sheet name and cell location to begin the analysis.
- Results are displayed for the provided cell and its dependencies.

## Output 📁
The tool generates output files containing extracted formulas, constants, and exceptions. Output files are saved in the _results_ directory as the script.
- ... \__formulas.txt_ file: contains all the real logic extracted from the excel cells. These are all the formulas/calculations that are needed to become the wanted result!
- ... \__values.txt_ file: contains all the values that where constants where found in the excel. This means that there was no real logic found behind these cells! We assume these values are (in case of Personal Income Tax - Tax Calculation (PB-berekening.xlsx)) display codes, or constants depending on region etc. E.g. the boundaries of the salary.
- ... \__exceptions.txt_file: contains all the edge cases of cell contents, which the script isn't build for. These cells need to be checked manually for anomallities or faults. Where each cell is used needs to be checked. A good example is e.g. an empty excel file!

## Contribution
This script was designed with a limitted amount of time available! So not all the possible features are included, a big scope is still not fully exploited! 

## License