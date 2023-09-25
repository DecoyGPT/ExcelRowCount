# Excel Row Counter

This Python script calculates the number of rows with data for each sheet in an Excel workbook and then outputs the grand total of rows with data across all sheets.

## Prerequisites

Before you begin, ensure you have met the following requirements:

- **Python**: You must have [Python](https://www.python.org/downloads/) installed on your machine.
- **openpyxl**: This script uses the `openpyxl` library to interact with Excel files. Install it via pip:

```bash
pip install openpyxl
```

## Usage

To use the Excel Row Counter script, follow these steps:

1. Clone this repository:
```bash
git clone https://github.com/DecoyGPT/ExcelRowCount.git
cd ExcelRowCount
```

2. Run the script by passing the path to your Excel file as an argument:
```bash
python ExcelRowCount.py path_to_your_workbook.xlsx
```

Replace `path_to_your_workbook.xlsx` with the path to your Excel file.

## Output

The script will output the number of rows with data for each sheet, followed by the grand total for the entire workbook:

```
Sheet1 = 1001 rows
Sheet2 = 100 rows
...
Total Rows: 1101 rows
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

This project uses the following license: [MIT](https://choosealicense.com/licenses/mit/).
