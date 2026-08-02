// Title: C# – Remove Duplicate Rows Based on Columns A & B with Aspose.Cells
// Description: Load an Excel workbook, define the used range, and call Cells.RemoveDuplicates with a header flag and column offsets {0,1} to delete rows that share the same values in columns A and B, then save the cleaned file.
// Keywords: Aspose.Cells RemoveDuplicates C# | delete duplicate rows Excel .NET | remove duplicate rows columns A B | clean Excel data Aspose.Cells | C# Excel duplicate removal example
// Common Searches: Aspose.Cells remove duplicate rows C# | How to delete rows with same values in columns A and B using Aspose.Cells | Remove duplicates Excel file .NET example | C# code for Excel duplicate row cleanup
// Developer Intent: Eliminate rows that have identical values in columns A and B while keeping the header row intact.
// Use Cases: Sanitizing sales logs where product‑date pairs (A‑B) must be unique. | Preparing a master report by stripping duplicate entries from merged Excel sources. | Automating nightly batch jobs that consolidate multiple workbooks and remove redundant rows.
// AI Prompts: Write C# code with Aspose.Cells to remove duplicate rows based on columns C and D, preserving the header. | Explain the purpose of each parameter in Cells.RemoveDuplicates for a .NET developer. | Show an alternative method to purge duplicates by importing worksheet data into a DataTable, filtering, and writing back to Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsDuplicateRemoval
{
    // Load an Excel workbook, define the used range, and call Cells.RemoveDuplicates with a header flag and column offsets {0,1} to delete rows that share the same values in columns A and B, then save the cleaned file.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the sheet
            int startRow = 0;                         // Include header row
            int startColumn = 0;                      // Column A (0‑based)
            int endRow = cells.MaxDataRow;            // Last row containing data
            int endColumn = cells.MaxDataColumn;      // Last column containing data

            // Remove duplicate rows based on values in columns A and B.
            // hasHeaders = true because the first row contains column titles.
            // columnOffsets = {0,1} specifies that columns A and B are used for duplicate comparison.
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { 0, 1 });

            // Save the cleaned workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
