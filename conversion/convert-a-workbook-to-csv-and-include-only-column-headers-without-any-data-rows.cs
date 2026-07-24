// Title: Create a Header‑Only CSV from an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, extracts the first row (column headers) from the first worksheet, copies those cells to a new workbook, and saves it as a CSV file that contains only the header row—no data rows are written.
// Keywords: Aspose.Cells | C# | CSV export | header only | Excel to CSV | extract column headers | save header row | convert workbook to CSV | Aspose.Cells CSV header | C# Excel header export
// Common Searches: Aspose.Cells export only header row to CSV | C# save Excel column names as CSV | How to generate CSV with just column headers using Aspose.Cells | Create header‑only CSV from .xlsx in .NET | Extract first row from Excel to CSV C#
// Developer Intent: Produce a CSV file that contains only the worksheet’s header row, omitting all data rows.
// Use Cases: Provide a template CSV for bulk data import where only column names are required. | Supply a schema file for downstream systems that consume CSV headers. | Generate a header‑only file for documentation or API contracts. | Create a lightweight CSV for validating column order without loading data.
// AI Prompts: Write C# code using Aspose.Cells to read an Excel file and save only the first row as a CSV. | Explain how to determine the last populated column in the header row with Aspose.Cells before exporting. | Give a step‑by‑step tutorial for creating a header‑only CSV from any workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderOnlyCsv
{
    // Loads an Excel file, extracts the first row (column headers) from the first worksheet, copies those cells to a new workbook, and saves it as a CSV file that contains only the header row—no data rows are written.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (any Excel format)
            string sourcePath = "input.xlsx";

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Access the first worksheet (assumed to contain the data)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Determine the last column that contains data in the header row (row 0)
            int lastColumn = sourceCells.MaxDataColumn;

            // Create a new workbook that will hold only the header row
            Workbook headerOnlyWorkbook = new Workbook();
            Worksheet headerSheet = headerOnlyWorkbook.Worksheets[0];
            Cells headerCells = headerSheet.Cells;

            // Copy each cell value from the header row of the source sheet to the new sheet
            for (int col = 0; col <= lastColumn; col++)
            {
                // Preserve the original value (string, number, etc.)
                headerCells[0, col].Value = sourceCells[0, col].Value;
            }

            // Save the new workbook as CSV. Only the header row will be written.
            string outputCsvPath = "headers_only.csv";
            headerOnlyWorkbook.Save(outputCsvPath, SaveFormat.Csv);

            Console.WriteLine($"Header-only CSV saved to: {outputCsvPath}");
        }
    }
}
