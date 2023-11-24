# Python-Data-Visualization-and-Summary-Tool

This Python program provides a graphical user interface (GUI) for loading and visualizing data from JSON or Excel files. It utilizes the Tkinter library for the GUI and integrates data visualization using Matplotlib. The tool generates summary statistics, visualizations (histograms, bar charts, scatter plots, and pie charts), and exports the summary report to an Excel file.

## Features

- **File Loading**: Load data from JSON or Excel files.
- **Summary Statistics**: Display summary statistics of the loaded data including mean, median, min, max, etc.
- **Data Visualizations**: Generate histograms, bar charts, scatter plots, and pie charts based on the loaded dataset.
- **Export to Excel**: Export the loaded data, summary statistics, and summary text content to an Excel file.

## Installation and Usage

### Prerequisites

- Python 3.x installed
- Required Python libraries: `tkinter`, `pandas`, `matplotlib`, `xlsxwriter`

### Installation

1. Clone or download the repository.
2. Install the required libraries if not already installed:
    ```
    pip install pandas matplotlib xlsxwriter
    ```

### Usage

Run the program by executing the `main.py` file:
```
python main.py
```

1. **Load Data**: Click on "File" -> "Load File" to load your data file (JSON or Excel).
2. **View Summary**: Summary statistics and visualizations will be displayed in the GUI.
3. **Export Summary**: Click on "File" -> "Export to Excel" to save the summary report in an Excel file.

## Notes

- Ensure your data file is in JSON or Excel format.
- For specific chart types or customizations, modify the code accordingly in the `load_data` function.
- Feel free to expand functionalities or improve the GUI according to your needs.
