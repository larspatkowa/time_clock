# Time Clock Calculator

## Overview
Simple GUI time clock calculator using Tkinter and readings from a CSV file

## Features
- **File Selection**: User can browse files to select input CSV file
- **Time Calculation**: Calculates total time clocked in for each user.
- **Abnormality Detection**: The script identifies and reports any abnormalities in the time clock records, such as unmatched 'IN' or 'OUT' entries.

## Usage
To use the time clock calculator, run the script and follow the prompts to select your CSV file.

## Dependencies
- Python 3.x
- Tkinter (included with standard Python distribution)

## Example Input CSV File Format
The Input CSV file should have the following columns:
- `User ID`
- `User Name`
- `Check Time`
- `Status` (either 'IN' or 'OUT')

Example:
```csv
User ID,User Name,Check Time,Status
1,Carol,2024/11/01 14:30:37,OUT
1,Carol,2024/11/01 08:15:00,IN
...