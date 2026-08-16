// Title: C# – Load a CSV with custom delimiter and header row into Aspose.Cells Workbook, then save as XLSX
// Description: Demonstrates using Aspose.Cells TxtLoadOptions in C# to import a CSV file with a specified separator, treat the first line as column headers, optionally convert numeric strings, access cells, and export the workbook to XLSX.
// Keywords: Aspose.Cells CSV import C# | TxtLoadOptions delimiter | HeaderRowsCount Aspose | convert numeric CSV Aspose.Cells | save workbook as XLSX | load semicolon delimited CSV | C# Excel file generation
// Common Searches: Aspose.Cells load CSV with custom separator C# | How to set header row when importing CSV in Aspose.Cells | Convert CSV numeric values to numbers using Aspose.Cells | Save imported CSV as XLSX with Aspose.Cells | C# TxtLoadOptions example for CSV
// Developer Intent: Import a CSV file using a chosen delimiter, map the first row to column headers, and write the result to an XLSX workbook with Aspose.Cells in C#.
// Use Cases: Transform a semicolon‑separated report that includes column names into a native Excel file for further analysis. | Preserve numeric data types when converting legacy CSV data to XLSX, enabling calculations and charting. | Programmatically validate or modify specific cells after loading a CSV before exporting to other formats.
// AI Prompts: Generate C# code to load a pipe‑delimited CSV with two header rows using Aspose.Cells. | Show how to export the workbook to PDF after setting the first row as headers and converting numeric strings. | Provide an example that reads the CSV into a DataTable, then writes it to an XLSX file with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates using Aspose.Cells TxtLoadOptions in C# to import a CSV file with a specified separator, treat the first line as column headers, optionally convert numeric strings, access cells, and export the workbook to XLSX.
class LoadCsvWithHeaders
{
    static void Main()
    {
        // Path to the CSV file
        string csvPath = "input.csv";

        // Configure CSV load options
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ';';          // specify the delimiter
        loadOptions.HeaderRowsCount = 1;      // treat the first row as header
        loadOptions.ConvertNumericData = true; // optional: convert numeric strings to numbers

        // Load the CSV into a workbook using the configured options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Example: read a value from the first data row (row 2, column A)
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("First data cell (A2): " + sheet.Cells["A2"].StringValue);

        // Save the workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
