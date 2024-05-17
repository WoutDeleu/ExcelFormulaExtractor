# Excel Formula Extractor 📊 📈 📉

This tool is designed to extract formulas, constants, and exceptions from an Excel workbook, specifically for the file _PB-berekeningen.xlsx_. It provides insights into the calculations made within the workbook, aiding in analysis and understanding. The tool extracts formulas and writes them in a more readable, standard syntax, all combined in one file.

The idea is to summarize and analyze complex and long Excel files without having to click through all the involved cells manually!

If you have any questions about contributing, functionality, etc., don't hesitate to contact me!

## Contents
- [Installation 🐍](#Installation-🐍)
- [How To Use](#How-to-Use)
    - [Without command line arguments](#Without-command-line-arguments)
    - [Command line arguments](#Command-line-arguments)
- [How it works 🧑‍🏫](#How-it-works-🧑‍🏫)
- [Output 📁](#Output-📁)
- [Contribution](#Contribution)


## Installation 🐍
1. Clone or download this repository, and enter the folder.
2. Ensure you have Python 3.x installed and Python added to the path of your device, as well as pip configured.
3. Install the required dependencies using the following command:
    ```shell
    pip install -r requirements.txt 
    ```
4. Run the script with the desired flags.

## How to Use
### Without command line arguments
1. Execute the script without flag options.
    ```shell
    python main.py
    ```
2. Select the Excel file you want to analyze (this can take some time with large files).
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
You can't provide just the --cell argument or just the --sheetname argument. Either both, or none, otherwise the program will fail. The same goes for the --full_analysis flag and the --single_cell flag!

## How it works 🧑‍🏫
### Flow
The first step in the flow of the program is handling the command line arguments and determining the flow of the program. 
- What file to analyze?
- Full Analysis / Single Cell?
- Starting Cell and sheet?

The core of the program is 3 stacks used to keep track of all the formulas, exceptions, and values. These store the items we want to extract from the program in a LIFO (_Last In First Out_) order. This is to recreate the dependency order of most programming languages, where the element that depends on a certain element is placed below it!

After that, every cell to analyze is resolved using the **resolve_cell** function. This means the content of the cell is analyzed. If it is an exception or a value, it is added to the corresponding list. If the cell contains a formula or some kind of logic, it is broken into parts.

The handle_formula function makes extensive use of the **extract_formulas** function. This is a _recursive_ function that breaks the formula into smaller parts and rebuilds it in the correct syntax. For example, ```SUM(A1:A4)``` is translated into ```A1+A2+A3+A4```. The second important functionality is that it extracts the involving cells and returns them to the **resolve_cell** call. This makes the function able to resolve and translate all the cells involved for one single cell! **extract_formulas** is a _recursive_ function, which means it calls itself. It breaks up the function into parts separated by brackets, if-statements, operators, etc., and extracts everything from these smaller parts.

### Supported Excel functions
The Excel functions that are handled by this script (some more thoroughly than others) are:
- _IFERROR_ (not extensive)
- _IF_
- _SUM_
- _VLOOKUP_
- _ROUND_
- _DATEDIF_ (not extensive)
- _DATE_ (not extensive)
- _LEFT_
- _MID_
- _RIGHT_

### Variable Naming in Results
When looking at the resulting files, you will notice a specific naming convention. Every variable name is built up of 2 parts: sheet name and cell number.
- **Sheet name:** referring to the sheet name of the original Excel file.
- **Cell number:** referring to the cell location on that sheet!

### Running a Full Analysis
- The tool iterates through predefined starting cells, extracting formulas, constants, and exceptions.
- Results are saved to separate files named according to the data type and cell position.

### Specific Analysis
- Specify the sheet name and cell location to begin the analysis.
- Results are displayed for the provided cell and its dependencies.

### Caution
- Manually check the order of the dependencies! There are some known bugs that occurred with the order of the formulas. This should be quite rare, but caution is needed when analyzing results.
- Look out for double calculated cells. At the moment, it is not possible to track all the already visited cells _across different runs_. Meaning that when a cell is used a lot across different cells and branches out extremely (like for example Tax Calculation - D113), this can cause exponential expansion of the result files. Some thorough analysis is therefore needed when encountering such cases, or preferably before running the tool. Tax Calculation - D113 for example is 'easy' to build manually, so if you leave this cell out of the calculations by the script, the complexity is decreased significantly!

## Output 📁
The tool generates output files containing extracted formulas, constants, and exceptions. Output files are saved in the _results_ directory as the script.
- `__formulas.txt` file: contains all the real logic extracted from the Excel cells. These are all the formulas/calculations that are needed to become the wanted result!
- `__values.txt` file: contains all the values that were constants found in the Excel. This means that there was no real logic found behind these cells! We assume these values are (in case of Personal Income Tax - Tax Calculation (PB-berekening.xlsx)) display codes or constants depending on region etc. E.g. the boundaries of the salary.
- `__exceptions.txt` file: contains all the edge cases of cell contents, which the script isn't built for. These cells need to be checked manually for anomalies or faults. Where each cell is used needs to be checked. A good example is e.g. an empty Excel file!

## Reporting Issues
If you encounter any bugs, have feature requests, or have questions about the project, please [open an issue](https://github.com/WoutDeleu/ExcelFormulaExtractor/issues) on GitHub. Be sure to provide as much detail as possible to help us understand and resolve the issue.

## Contribution
This script was designed with a limited amount of time available! So not all the possible features are included, a big scope is still not fully exploited! Feel free to contribute!

To help you get started, here is a list of current tasks and features we are working on.
- Refactoring and adding Docstrings to functions
    - Remove global variables
- Remove hard coded list of cells to follow with full analysis, and replace this with an extra command line argument
- _Order bug_:  In some cases is the order not correct! This can be due to circular dependencies
- Implement a large tracker for the full analysis, so no duplicated cells are being resolved across analyses!
- More extensive tests to be added
- Enhance handled functions within the application
    - Array Functions: Expand the array functionalities.
    - Date Functions: Implement comprehensive date functions.
    - Datedif: Add functionality to calculate the difference between dates.
    - Left/Right/Mid: Implement string manipulation functions.
- A language parser
- The TODO's inside the code
- A language parser to immediately get code in the correct syntax of a chosen language. E.g. running the code, and getting a fully functional Python file as a result

- Simplify formulas
    - Substituate formulas in one another
        - Direct assignments -> substituting
    - VLOOKUPs

### How to Contribute
We welcome contributions from the community! To ensure a smooth process, please follow the guidelines below:

1. **Fork the Repository**  
   Create a personal fork of the project on GitHub.

2. **Clone Your Fork**  
   Clone your fork to your local machine:
   ```bash
   git clone https://github.com/WoutDeleu/ExcelFormulaExtractor
    ```

3. **Create a Branch**
    Create a new branch for your changes:
    ```bash
    git checkout -b feature-branch
    ```

4. **Make Changes**
    Implement your changes, following the project's coding style and conventions. Ensure you write docstrings and comments where necessary.

5. **Write Tests**
    If your changes include new features or bug fixes, please write tests to verify your work. Place your tests in the appropriate directory.

<!-- 6. **Run Tests**
    Ensure all tests pass before submitting your contribution:
    ```bash
    pytest

    ```
7. **Commit and Push**
    Commit your changes with a descriptive message and push them to your fork:
    ```bash
    git add .
    git commit -m "Description of changes"
    git push origin feature-branch
    ``` -->

6. **Create a Pull Request**
    Go to the original repository and create a pull request from your fork. Provide a clear and detailed description of your changes.