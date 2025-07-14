# ASN-remapper

**ASN Template Mapper with Lookup** is a powerful desktop tool (Python + Tkinter) that helps you remap, combine, and transform columns from spreadsheets (Excel/CSV) into a standardized ASN (Advanced Shipping Notice) template.  
It supports multi-select, direct mapping, manual input, and reference lookups, making it ideal for warehousing, logistics, and supply chain data processing.

## Features

- **Intuitive GUI:** No coding required – select files and map columns visually.
- **Flexible Mappings:**  
  - **Direct Mapping:** Map source columns directly to ASN template fields.
  - **Multi-Select:** Combine multiple columns into one ASN field (with ordered selection).
  - **Manual Input:** Enter static values for ASN fields.
  - **Lookup Mapping:** Transform values using external reference files.
- **Lookup Management:** Add and manage multiple lookup/reference files easily.
- **Column Requirements:** Required ASN fields are clearly indicated.
- **Output:** Generates an ASN template as Excel file, with a summary of mappings.

## Installation

1. **Clone the Repository:**
    ```bash
    git clone https://github.com/abamakbar07/ASN-remapper.git
    cd ASN-remapper
    ```

2. **Install Dependencies:**
    ```bash
    pip install pandas openpyxl xlrd pyxlsb
    ```

    > **Note:** Tkinter is included with standard Python distributions.

## Usage

1. **Run the App:**
    ```bash
    python script.py
    ```

2. **Workflow:**
    - **Select Source File:** Choose your Excel/CSV file.
    - **(Optional) Add Lookup Reference Files:** Load reference files for value mapping.
    - **Map Columns:** For each required ASN field, choose mapping type and configure.
    - **Generate ASN Template:** Output Excel file will be saved in the current directory.

3. **Mapping Types:**
    - **Direct:** Choose a source column.
    - **Multi-Select:** Combine several columns (in order) using `|` separator.
    - **Lookup:** Map source values using external lookup file.
    - **Manual Input:** Enter a fixed/static value.
// 
// ## Screenshots
// 
// *(You can add screenshots of the GUI here)*

## Example

Suppose you have a source file `shipment.csv` and a lookup file `owners.xlsx`:

- Map the "Owner" ASN field using a lookup from `owners.xlsx`.
- Map "LOTTABLE01" by combining "Batch" and "Lot" columns.
- Fill "Hold Code:" with a static value.

The output will be an Excel file with all required ASN columns properly mapped.

## Troubleshooting

- **File Loading Issues:** Make sure your source and lookup files are in supported formats (`.xlsx`, `.xls`, `.csv`, `.xlsb`).
- **Missing Columns:** Only required ASN columns are shown by default.
- **Dependencies:** If you see errors about missing packages, run `pip install -r requirements.txt`.

## License

MIT

## Credits

Developed by [abamakbar07](https://github.com/abamakbar07).
