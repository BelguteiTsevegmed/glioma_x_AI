# Glioma-X-AI Data Entry Automation

This repository contains a Python script designed to automate part of my data entry task in a research project investigating a correlation between radiological images of primary CNS neoplasms and their WHO classification using machine learning. 

The script reads data from an Excel workbook, parses specific gene information from a cell, and adds this information to corresponding cells. It uses the openpyxl library to read and manipulate Excel workbooks.

## Features

- Read data from an Excel workbook.
- Parse gene information from a specified cell.
- Add gene presence information to corresponding cells.

## Dependencies

- Python 3.7 or newer
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Installation
Ensure you have Python 3.7 or newer installed. You can download Python from the [official website](https://www.python.org/downloads/).

Install openpyxl with pip:

```sh
pip install openpyxl
```

## Usage
Before running the script, make sure you have an Excel workbook with the expected format in the same directory as the script. The workbook should be named 'Glioma-XAI-kopia.xlsx'.

To run the script:

```sh
python glioma_xai.py
```

After the script finishes, it will save the updated data to 'Glioma-XAI-kopia.xlsx'. Cells that the script has not assigned a value to will be filled with '-'.

## Note
This script is customized to my specific data and needs for the mentioned research project. It's not intended for general use. However, you might find some parts of the code or the approach beneficial for similar tasks.

## License
This project is licensed under the MIT License.
