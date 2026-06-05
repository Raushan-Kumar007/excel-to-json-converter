# Excel to JSON Converter

This is a lightweight Streamlit web application that allows you to extract hierarchical product relationships from Excel files and generate a structured Excel output.

## Features
- **Multi-File Upload:** Upload one or more `.xlsx` or `.xls` files at the same time. The application automatically combines the data across all uploaded files.
- **Auto-Detection:** Automatically maps parent-child product relationships based on column names (e.g., Parent ID, Child ID, Parent Name, Child Name).
- **Hierarchical Tree Construction:** Builds a fully nested JSON structure representing your products and all of their sub-components.
- **Excel Export:** Generates a clean Excel file containing the extracted `Product Name`, `Id`, and the full `Product Config Description` (the JSON block) for each top-level parent product.

## Installation

1. Ensure you have Python installed.
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the App

To start the application locally, run the following command in your terminal:
```bash
streamlit run app.py
```

The application will open automatically in your default web browser at `http://localhost:8501`.

## Usage
1. Drag and drop your Excel files into the upload area.
2. Review the combined Data Preview to ensure everything was loaded correctly.
3. Click on the **"Click to Preview Generated Excel"** expander to see exactly what the final output will look like.
4. Click the **"Download Excel"** button to save the final `.xlsx` file to your computer.
