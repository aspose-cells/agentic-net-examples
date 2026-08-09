// Title: Export Excel Headers to CSV with Aspose.Cells (C#) – Header‑Only File
// Description: Loads an Excel workbook, copies the first row (the column headers) into a new workbook, and saves that workbook as a CSV using TxtSaveOptions. The resulting file contains only the header line, no data rows.
// Keywords: Aspose.Cells | C# | CSV export | header only | Excel to CSV | TxtSaveOptions | .NET workbook conversion | extract column headers
// Common Searches: Aspose.Cells export only header row to CSV | C# save Excel header as CSV | how to create CSV with just column names using Aspose | trim blank rows when saving CSV Aspose.Cells | convert Excel to CSV without data rows .NET
// Developer Intent: Create a CSV file that includes only the column names from an Excel worksheet.
// Use Cases: Generate template CSV files that list required fields for data import pipelines. | Provide documentation‑ready column listings without exposing any record data. | Supply API schema definitions that need only the header row from an Excel source.
// AI Prompts: Write C# code with Aspose.Cells that extracts the first row of an Excel sheet and saves it as a CSV containing only the headers. | Create a reusable method that takes an Excel file path and returns a CSV string with just the column headers, using TxtSaveOptions. | Explain how TxtSaveOptions can be configured to trim leading blank rows and columns when exporting a header‑only CSV with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // For TxtSaveOptions

// Loads an Excel workbook, copies the first row (the column headers) into a new workbook, and saves that workbook as a CSV using TxtSaveOptions. The resulting file contains only the header line, no data rows.
class WorkbookToCsvHeadersOnly
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string sourcePath = "input.xlsx";
        Workbook sourceWb = new Workbook(sourcePath);

        // Create a new workbook that will contain only the header row
        Workbook headerOnlyWb = new Workbook();
        Worksheet srcSheet = sourceWb.Worksheets[0];
        Worksheet dstSheet = headerOnlyWb.Worksheets[0];

        // Determine the last column that contains data in the header row
        int lastCol = srcSheet.Cells.MaxDataColumn;

        // Copy each header cell from the first row (row index 0) to the new workbook
        for (int col = 0; col <= lastCol; col++)
        {
            // Read the header value from the source sheet
            string headerValue = srcSheet.Cells[0, col].StringValue;

            // Write the header value to the destination sheet
            dstSheet.Cells[0, col].PutValue(headerValue);
        }

        // Configure CSV (text) save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            // Export only the active sheet (default), no need to set ExportAllSheets
            // Trim leading blank rows/columns to mimic Excel's behavior
            TrimLeadingBlankRowAndColumn = true
        };

        // Save the header‑only workbook as CSV
        string outputPath = "headers_only.csv";
        headerOnlyWb.Save(outputPath, csvOptions);

        Console.WriteLine($"CSV file with only column headers saved to: {outputPath}");
    }
}
