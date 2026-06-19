# STATCAL ONLINE - Panel Mean Line Chart Builder

A Python Streamlit application for creating publication-ready mean line charts from panel datasets.

## Features

- Upload Excel data (.xlsx / .xls)
- Training data link in the main page and sidebar
- Flexible numeric variable selection
- Mean line chart based on selected numeric variables
- Split lines by categorical variables such as Company or Year
- Multi-panel charts by numeric variables
- Publication-ready background themes and color palettes
- Dot markers and mean value labels on the line chart
- High-resolution PNG export: 300, 600, 900, 1200, and 1500 DPI
- Descriptive statistics table: minimum, maximum, mean, and standard deviation
- Excel export for descriptive statistics and chart data

## Run

```bash
pip install -r requirements.txt
python -m streamlit run main.py
```
