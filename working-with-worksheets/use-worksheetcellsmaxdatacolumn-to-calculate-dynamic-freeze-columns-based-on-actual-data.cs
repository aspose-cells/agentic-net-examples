// Title: Dynamic Column Freeze in Aspose.Cells .NET Using Worksheet.Cells.MaxDataColumn
// Description: Demonstrates how to create a workbook, populate it with sample data, retrieve the zero‑based index of the last populated column via Worksheet.Cells.MaxDataColumn, and apply FreezePanes to lock all columns up to that point before saving the file.
// Keywords: Aspose.Cells MaxDataColumn | dynamic column freeze .NET | FreezePanes based on data | last populated column Aspose | C# spreadsheet freeze panes | auto freeze columns Aspose.Cells
// Common Searches: freeze columns up to last data column Aspose.Cells | Worksheet.Cells.MaxDataColumn example C# | how to set FreezePanes dynamically in .NET | auto‑detect last column and freeze in Aspose.Cells
// Developer Intent: Detect the final data column and freeze all preceding columns programmatically.
// Use Cases: Generate reports where column count varies and headers must stay visible. | Export spreadsheets with unknown width while keeping data columns locked for review. | Prepare workbooks for printing where all populated columns should remain static during horizontal scrolling.
// AI Prompts: Write C# code that uses Worksheet.Cells.MaxDataColumn to freeze columns and also freezes the top row for header visibility. | Explain the relationship between MaxDataColumn and FreezePanes and show how to implement a dynamic freeze pane in Aspose.Cells. | Provide a snippet that opens an existing Excel file, finds the maximum data column, and applies FreezePanes to keep all data columns in view while scrolling.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicFreeze
{
    // Demonstrates how to create a workbook, populate it with sample data, retrieve the zero‑based index of the last populated column via Worksheet.Cells.MaxDataColumn, and apply FreezePanes to lock all columns up to that point before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to demonstrate dynamic freeze based on actual data
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Score");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(85);
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(92);
            // Add data in a later column to test dynamic detection
            worksheet.Cells["E1"].PutValue("Remarks");
            worksheet.Cells["E2"].PutValue("Excellent");
            worksheet.Cells["E3"].PutValue("Good");

            // Determine the maximum column index that contains data
            int maxDataColumn = worksheet.Cells.MaxDataColumn; // zero‑based index

            // If there is data, freeze all columns up to the last data column
            if (maxDataColumn >= 0)
            {
                // Freeze panes at the cell just after the last data column.
                // Row index = 0 (top of sheet), column index = maxDataColumn + 1.
                // freezedRows = 0 (no frozen rows), freezedColumns = maxDataColumn + 1.
                worksheet.FreezePanes(0, maxDataColumn + 1, 0, maxDataColumn + 1);
            }

            // Save the workbook
            workbook.Save("DynamicFreezeColumns.xlsx");
        }
    }
}
