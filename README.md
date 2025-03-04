# DFMEA Parser

A Python script for parsing and processing DFMEA (Design Failure Mode and Effects Analysis) Excel files.

## Features

- Splits rows based on semicolon-separated values
- Handles multiple entries in Recommended Actions column
- Manages Comments column entries
- Interactive replacement of values
- Preserves original data while creating processed versions

## Requirements

- Python 3.x
- pandas
- numpy
- openpyxl

## Installation

1. Clone the repository:
```bash
git clone https://github.com/foadkrmn/DFMEA-parser.git
cd DFMEA-parser
```

2. Create and activate a virtual environment:
```bash
python3 -m venv .venv
source .venv/bin/activate  # On macOS/Linux
```

3. Install required packages:
```bash
pip install pandas numpy openpyxl
```

## Usage

1. Place your Excel file in the same directory as the script
2. Run the script:
```bash
python DFMEA_parser
```
3. Follow the prompts to:
   - Enter the Excel file name
   - Specify the Recommended Actions column name
   - Specify the Comments column name
   - Review and modify entries as needed

## Output

The script generates:
- A processed Excel file with split rows
- Updated entries based on user modifications
- Preserved original data in separate columns 