// Title: C# – Split a CSV address column into Street, City, and Zip with Aspose.Cells
// Description: Loads a CSV into a workbook, uses Aspose.Cells TextToColumns to separate a combined address field (comma‑delimited) into three columns, and saves the result as an XLSX file.
// Keywords: Aspose.Cells CSV address split | C# TextToColumns | split address column | CSV to Excel conversion | extract street city zip | .NET spreadsheet library
// Common Searches: Aspose.Cells split address column C# | TextToColumns example for CSV | how to separate street city zip in Excel using code | C# import CSV and parse address fields
// Developer Intent: Separate a single address column from a CSV into distinct street, city, and zip columns and export to Excel.
// Use Cases: Transform raw mailing‑list CSVs into structured Excel sheets for mail‑merge. | Prepare customer data for reporting by breaking out address components. | Cleanse and normalize address information before loading into a database.
// AI Prompts: Show how to change the delimiter to a semicolon for address splitting. | Add code that trims whitespace from each resulting address part after TextToColumns. | Explain handling rows with missing zip codes while using TextToColumns.

using System;
using Aspose.Cells;

namespace AddressSplitExample
{
    // Loads a CSV into a workbook, uses Aspose.Cells TextToColumns to separate a combined address field (comma‑delimited) into three columns, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Load CSV data (comma‑separated) starting at cell A1 (row 0, column 0)
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = ',';               // CSV delimiter
            cells.ImportCSV(csvPath, loadOptions, 0, 0);

            // Determine the number of rows that contain data
            int totalRows = cells.MaxDataRow + 1;       // zero‑based index + 1

            // Index of the column that holds the combined address (e.g., column A => index 0)
            int addressColumnIndex = 0;

            // Prepare split options: split the address by comma into separate columns
            TxtLoadOptions splitOptions = new TxtLoadOptions();
            splitOptions.Separator = ',';               // Address delimiter

            // Perform the split – the original column will be expanded into new columns
            cells.TextToColumns(0, addressColumnIndex, totalRows, splitOptions);

            // Save the result to an Excel file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
