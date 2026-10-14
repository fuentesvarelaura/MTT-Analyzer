# MTT Analyzer

MTT Analyzer is a macOS desktop application designed to simplify the analysis of MTT assay results from Excel files containing 96-well plate reader measurements.

## Features

- Import Excel files (`.xlsx`).
- Interactive visualization of a 96-well plate (8 rows × 12 columns).
- Assign wells to different experimental conditions using color coding.
- Customize condition names and colors.
- Designate a condition as the control.
- Calculate the control mean and apply background noise correction.
- Calculate percentages relative to the corrected control.
- Write the results directly into the Excel file.

## How It Works

### 1. Select the Excel File

Select the `.xlsx` file containing the MTT assay measurements.

The application reads the plate data from rows 27–34 and columns C–N of the active worksheet.

### 2. Define Experimental Conditions

Create experimental conditions, assign a name and color to each one, and select the corresponding wells using the graphical plate interface.

Each condition is displayed in a different color for easy identification.

### 3. Select the Control

Designate one of the conditions as the control and assign its corresponding wells.

The application calculates the mean absorbance of the control wells and subtracts 0.04 to correct for background noise.

### 4. Calculate Relative Percentages

For each well belonging to a non-control condition, the application calculates the percentage relative to the corrected control using the following formula:

\[
\text{Relative percentage} =
\frac{\text{Well absorbance} - 0.04}
{\text{Control mean} - 0.04}
\times 100
\]

### 5. Export the Results

The application writes the calculated results, condition labels, and condition averages into the original Excel file.

**Important:** The original Excel file is overwritten. Always create a backup copy before running the analysis.

## Requirements

- macOS.
- Python 3.
- The Python package `openpyxl`.

## Installation

Install the required dependencies by running the following command in the terminal from the project directory:

```bash
python -m pip install -r requirements.txt
```

## Technologies Used

- **Python** — application logic.
- **Tkinter** — graphical user interface.
- **Openpyxl** — reading and writing Excel files.
- **PyInstaller** — packaging the application as a desktop app.

## Notes and Limitations

- Plate measurements must be located in the range C27:N34 of the active worksheet.
- The background correction value is fixed at 0.04.
- Results are written directly into the selected Excel file.
- Always verify the results and keep a backup of the original experimental data.
