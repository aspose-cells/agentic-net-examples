// Title: C# – Conditionally Freeze Top Row in Worksheets with Over 100 Rows Using Aspose.Cells
// Description: Loads a workbook, checks each worksheet’s data row count with MaxDataRow, and applies FreezePanes to the first row only when the sheet contains more than 100 rows, then saves the file.
// Keywords: Aspose.Cells | C# | .NET Excel | FreezePanes | conditional freeze panes | MaxDataRow | worksheet row count | freeze header row | large worksheets | Excel automation
// Common Searches: Aspose.Cells freeze first row if rows > 100 | C# conditional FreezePanes based on row count | How to use MaxDataRow with FreezePanes in Aspose.Cells | Iterate all worksheets and apply FreezePanes .NET | Freeze top row for large Excel sheets using Aspose
// Developer Intent: Apply a freeze pane to the header row of any worksheet that contains more than 100 data rows.
// Use Cases: Keep column headers visible while scrolling through extensive reports generated with Aspose.Cells. | Standardize the view of exported spreadsheets that exceed a set row threshold. | Prepare workbooks for printing or sharing, ensuring the header stays in view on large sheets.
// AI Prompts: Generate C# code with Aspose.Cells that freezes the top row only when a worksheet has over 100 data rows. | Show how to extend the example to also freeze the first column when a sheet exceeds a given column count. | Provide a snippet that logs the names of worksheets where FreezePanes was applied. | Create a version that uses a configurable row threshold instead of a hard‑coded 100 rows. | Explain how to handle empty worksheets when using MaxDataRow and FreezePanes.

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // Loads a workbook, checks each worksheet’s data row count with MaxDataRow, and applies FreezePanes to the first row only when the sheet contains more than 100 rows, then saves the file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (lifecycle: load)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last row that contains data (zero‑based index)
                // MaxDataRow returns -1 if the sheet is empty
                int lastDataRow = sheet.Cells.MaxDataRow;

                // Calculate total number of rows with data
                int totalRows = lastDataRow + 1; // convert to 1‑based count

                // If the worksheet has more than 100 rows, apply freeze panes
                if (totalRows > 100)
                {
                    // Freeze the first row (row index 1 means the freeze line is after row 0)
                    // Parameters: row, column, freezedRows, freezedColumns
                    sheet.FreezePanes(1, 0, 1, 0);
                }
            }

            // Save the modified workbook (lifecycle: save)
            workbook.Save("output.xlsx");
        }
    }
}
