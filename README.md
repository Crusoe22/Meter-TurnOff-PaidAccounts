# Meter-TurnOff-PaidAccounts

# Data Export and Formatting Tool

This tool is designed to export data from an ArcGIS feature class to an Excel spreadsheet and then format the spreadsheet according to specified requirements.

## Requirements
- Python 3.x
- ArcPy
- Pandas
- Openpyxl

## Setup
Ensure that ArcPy and its dependencies are properly installed. Additionally, ensure that the Python environment path is correctly set to: `C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3`.

## Usage

### `grabdata()`
- Functionality: Imports data from the feature class `HUD_LGIM.dbo.NIGHTDUTYACCOUNTS`.
- Input:
  - Feature class path (`fc`)
  - Output Excel file path (`output_excel`)
  - List of fields to extract from the feature class (`fields`)
- Output:
  - Saves the extracted data into the specified Excel file.

### `formatexcel()`
- Functionality: Formats the Excel sheet.
- Input:
  - Input Excel file path (`input_excel`)
- Output:
  - Saves the formatted Excel file.

### `color_width()`
- Functionality: Sets column width and applies color fill to specific cells based on data conditions.
- Input:
  - None
- Output:
  - Saves the updated Excel file.

## Instructions
1. Run the `grabdata()` function to import data from the feature class and save it to an Excel file.
2. Run the `formatexcel()` function to format the Excel sheet.
3. Run the `color_width()` function to set column width and apply color fills based on data conditions.

## Example
```python
import arcpy
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

grabdata()
formatexcel()
color_width()
