// Title: C# – Load CSV, Remove Duplicate Rows by Key Column, and Export to XLSX with Aspose.Cells
// Description: A concise example that shows how to import a CSV file into an Aspose.Cells workbook, identify and delete rows that share the same value in a specified column using the RemoveDuplicates method, and save the cleaned data as an XLSX file.
// Keywords: Aspose.Cells C# | CSV to Excel conversion | remove duplicate rows | RemoveDuplicates method | key column duplicate detection | import CSV Aspose.Cells | export XLSX .NET | duplicate record elimination | Aspose.Cells .NET API
// Common Searches: Aspose.Cells remove duplicate rows C# | How to delete duplicate CSV records with Aspose.Cells | Convert CSV to XLSX and filter duplicates in .NET | RemoveDuplicates example Aspose.Cells | C# code to clean CSV data before Excel export
// Developer Intent: Import a CSV file, purge rows that duplicate a chosen column value, and write the result to an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Clean a product catalog CSV by removing duplicate SKU entries before generating an Excel price list. | Deduplicate a user‑email CSV export to ensure each address appears only once in the final report. | Process log files exported as CSV, keeping only the first occurrence of each timestamp for analysis in Excel.
// AI Prompts: Generate C# code with Aspose.Cells that loads a CSV, removes duplicate rows based on column index 2, and saves the output as XLSX. | Explain the purpose of each parameter in Cells.RemoveDuplicates when the source CSV contains a header row. | Show how to configure RemoveDuplicates to consider multiple columns (e.g., ID and Date) as the composite key in a .NET application.

using System;
using Aspose.Cells;

// A concise example that shows how to import a CSV file into an Aspose.Cells workbook, identify and delete rows that share the same value in a specified column using the RemoveDuplicates method, and save the cleaned data as an XLSX file.
class RemoveCsvDuplicates
{
    static void Main()
    {
        // Input CSV file path and output Excel file path
        string inputCsvPath = "input.csv";
        string outputExcelPath = "output.xlsx";

        // Create a new empty workbook
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as the delimiter and converting numeric strings to numbers
        cells.ImportCSV(inputCsvPath, ",", true, 0, 0);

        // Determine the used range after import
        int startRow = 0;                         // first row (including header)
        int startColumn = 0;                      // first column
        int endRow = cells.MaxDataRow;            // last row with data
        int endColumn = cells.MaxDataColumn;      // last column with data

        // Index of the key column used to identify duplicates (0‑based)
        // Change this value to the appropriate column index in your CSV
        int keyColumnIndex = 0;

        // Remove duplicate rows based on the key column.
        // The 'true' flag indicates that the first row contains headers.
        // The columnOffsets array specifies which columns are considered for duplicate detection.
        cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { keyColumnIndex });

        // Save the cleaned workbook as an XLSX file
        workbook.Save(outputExcelPath, SaveFormat.Xlsx);
    }
}
