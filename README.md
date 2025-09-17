# stratFISCHER

**stratFISCHER** is a Python-based tool for generating Fischer plots directly from stratigraphic data.  
It automates data extraction, processing, and plotting, producing both graphical outputs and Excel reports.  
This tool was developed as part of the manuscript submitted to *Computers & Geosciences*.  

## Features
- Reads stratigraphic cycle thickness data from Excel files (.xls format).  
- Calculates cycle thickness, mean thickness, and cumulative departure from mean thickness (CDMT).  
- Produces two Fischer plots:  
  1. Thickness vs. CDMT  
  2. Cycle Number vs. CDMT  
- Exports processed data and figures into a new Excel workbook.  

## Requirements
- Python 3.x  
- Libraries:  
  - numpy  
  - matplotlib  
  - xlrd  
  - xlsxwriter  
- Standard Python library: os  

## Installation
1. Download and install [Python](https://www.python.org/downloads/).  
2. Install the required libraries by running:  
   ```bash
   pip install numpy matplotlib xlrd xlsxwriter
   ```

## Usage
1. Place your stratigraphic Excel data in the same folder as `stratFISCHER.py`.  
2. Run the script by double-clicking `stratFISCHER.py` or executing:  
   ```bash
   python stratFISCHER.py
   ```
3. Processed Excel files will be generated automatically in the same folder.  
   These will include:  
   - A processed data table  
   - Fischer plot of Thickness vs. CDMT  
   - Fischer plot of Cycle Number vs. CDMT  

## Example Dataset
A sample dataset is included in this repository:  
- `example_stratigraphic_data.xlsx` (50 cycles of randomly generated thickness values).  

You can test the program using this dataset:  
1. Ensure `example_stratigraphic_data.xlsx` is in the same folder as `stratFISCHER.py`.  
2. Run the program.  
3. A new Excel file (`processed_example_stratigraphic_data.xlsx`) will be created with plots and data outputs.  

## License
This project is licensed under the MIT License.  
