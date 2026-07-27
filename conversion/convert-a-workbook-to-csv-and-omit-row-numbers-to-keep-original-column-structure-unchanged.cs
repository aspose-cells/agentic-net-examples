// Title: C# – Convert Excel Workbook to CSV without Row Numbers using Aspose.Cells
// Description: Load any Excel workbook with Aspose.Cells, configure TxtSaveOptions for CSV, set TrimLeadingBlankRowAndColumn = false to keep blank rows/columns, and save a CSV that preserves the original column layout while omitting the automatic row‑number column.
// Keywords: Aspose.Cells CSV export C# | TxtSaveOptions TrimLeadingBlankRowAndColumn | remove row numbers CSV Aspose | preserve column structure CSV | Excel to CSV conversion .NET | save workbook as CSV without index
// Common Searches: Aspose.Cells save as CSV without row numbers | C# export Excel to CSV keep blank rows | TrimLeadingBlankRowAndColumn false CSV | prevent row index column in CSV Aspose | convert .xlsx to .csv preserving column order
// Developer Intent: Generate a CSV file from an Excel workbook in C# while retaining the exact column arrangement and avoiding any added row‑number column.
// Use Cases: Export financial or reporting spreadsheets to CSV for downstream systems that require precise column positions. | Batch‑process multiple workbooks into CSV files without altering blank rows or columns, ensuring data alignment. | Create CSV feeds for APIs or data pipelines where an extra row index would break the format.
// AI Prompts: Write C# code that uses Aspose.Cells to convert an .xlsx file to CSV, disables TrimLeadingBlankRowAndColumn, and removes the row‑number column. | Explain the impact of the TrimLeadingBlankRowAndColumn property on CSV output and how to keep the original layout. | Provide a reusable method that accepts input and output paths and converts any supported workbook to CSV without adding row numbers.

using System;
using Aspose.Cells;

// Load any Excel workbook with Aspose.Cells, configure TxtSaveOptions for CSV, set TrimLeadingBlankRowAndColumn = false to keep blank rows/columns, and save a CSV that preserves the original column layout while omitting the automatic row‑number column.
class WorkbookToCsv
{
    static void Main()
    {
        // Path to the source workbook (any supported format)
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file
        string csvPath = "output.csv";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options:
        // - TrimLeadingBlankRowAndColumn = false ensures that leading blank rows/columns
        //   are not removed, preserving the original column structure.
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.TrimLeadingBlankRowAndColumn = false;

        // Save the workbook as CSV without adding row numbers
        workbook.Save(csvPath, csvOptions);
    }
}
