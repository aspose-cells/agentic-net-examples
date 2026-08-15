// Title: C# – Freeze Panes on Worksheets with Over 10 Columns Using Aspose.Cells
// Description: Creates a workbook, adds worksheets with varying column counts, checks each sheet's MaxDataColumn, and applies FreezePanes(1,1,1,1) to lock the first row and column when the sheet contains more than ten populated columns, then saves the file.
// Keywords: Aspose.Cells | C# FreezePanes | MaxDataColumn | conditional freeze panes | Excel automation .NET | freeze first row column | wide worksheet handling
// Common Searches: Aspose.Cells freeze panes based on column count | C# freeze first row and column if more than 10 columns | How to use MaxDataColumn with FreezePanes in Aspose.Cells | Conditional FreezePanes example .NET | Freeze panes for large Excel sheets programmatically
// Developer Intent: Automatically lock the header row and first column on any worksheet that has more than ten data columns.
// Use Cases: Generate multi‑sheet reports where wide tables need frozen headers for easier navigation | Standardize freeze settings across all sheets in an automated Excel export | Improve readability of dashboards by preventing horizontal scrolling beyond a column threshold
// AI Prompts: Generate C# code using Aspose.Cells that freezes the top row and left column only when a worksheet contains more than a specified number of columns. | Describe how MaxDataColumn can be leveraged to decide when to call FreezePanes in an Aspose.Cells workbook. | Suggest alternative approaches to apply FreezePanes conditionally without looping through each worksheet.

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // Creates a workbook, adds worksheets with varying column counts, checks each sheet's MaxDataColumn, and applies FreezePanes(1,1,1,1) to lock the first row and column when the sheet contains more than ten populated columns, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Example: populate worksheets with sample data
            // Worksheet 0 will have 12 columns (will be frozen)
            Worksheet ws0 = workbook.Worksheets[0];
            for (int col = 0; col < 12; col++)
            {
                ws0.Cells[0, col].PutValue($"Header {col + 1}");
                ws0.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Worksheet 1 will have 8 columns (will not be frozen)
            Worksheet ws1 = workbook.Worksheets.Add("SmallSheet");
            for (int col = 0; col < 8; col++)
            {
                ws1.Cells[0, col].PutValue($"H{col + 1}");
                ws1.Cells[1, col].PutValue($"D{col + 1}");
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the number of columns that contain data
                // MaxDataColumn returns zero‑based index of the last column with data
                int lastDataColumnIndex = sheet.Cells.MaxDataColumn;

                // If there are more than 10 columns (i.e., index >= 10)
                if (lastDataColumnIndex >= 10)
                {
                    // Freeze the first row and first column (cell B2 is the freeze point)
                    // Parameters: row, column, freezedRows, freezedColumns
                    sheet.FreezePanes(1, 1, 1, 1);
                }
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("FreezePanesResult.xlsx");
        }
    }
}
