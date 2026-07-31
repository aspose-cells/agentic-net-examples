// Title: C# – Get Last Used Column and Freeze All Columns to Its Left with Aspose.Cells
// Description: This example creates a workbook, uses Cells.MaxDataColumn to find the index of the right‑most column that contains data, converts it to a 1‑based count, and then calls Worksheet.FreezePanes to lock every column up to that count while leaving rows unfrozen. The file is saved as FreezeColumnsResult.xlsx.
// Keywords: Aspose.Cells C# max data column | Worksheet FreezePanes column count | last used column Aspose.Cells | .NET Excel freeze columns programmatically | retrieve column count Aspose.Cells | freeze panes based on data range
// Common Searches: Aspose.Cells how to find last populated column in C# | freeze all columns up to last data column Aspose.Cells | Worksheet.FreezePanes example using MaxDataColumn | C# get column count and freeze panes in Excel
// Developer Intent: Identify the number of columns that contain data and apply a freeze pane that locks every column to the left of that boundary.
// Use Cases: Automatically lock header and data columns in generated reports so users can scroll horizontally without losing context. | Create a reusable utility that adapts freeze panes to worksheets of varying widths. | Prevent accidental edits to key columns in export templates by programmatically applying column freezes.
// AI Prompts: Generate C# code with Aspose.Cells that determines the last populated column and freezes all preceding columns. | Write a method named FreezeToLastColumn that returns the max data column count and applies Worksheet.FreezePanes accordingly. | Explain the relationship between Cells.MaxDataColumn and Worksheet.FreezePanes when locking columns based on actual data.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezeColumnsDemo
{
    // This example creates a workbook, uses Cells.MaxDataColumn to find the index of the right‑most column that contains data, converts it to a 1‑based count, and then calls Worksheet.FreezePanes to lock every column up to that count while leaving rows unfrozen. The file is saved as FreezeColumnsResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data – you can replace this with your own data loading logic
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["C1"].PutValue("Header3");
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].PutValue(30);

            // Retrieve the maximum data column index (zero‑based)
            int maxDataColumnIndex = cells.MaxDataColumn; // -1 if no data
            // Convert to column count (1‑based). If no data, count is 0.
            int maxDataColumnCount = maxDataColumnIndex >= 0 ? maxDataColumnIndex + 1 : 0;

            Console.WriteLine($"Maximum data column count: {maxDataColumnCount}");

            // Freeze all columns to the left of the last data column.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // We freeze rows = 0 (no row freeze) and columns = maxDataColumnCount.
            // The freeze position is the first cell after the frozen area,
            // therefore column parameter = maxDataColumnCount (zero‑based index of that cell).
            if (maxDataColumnCount > 0)
            {
                worksheet.FreezePanes(0, maxDataColumnCount, 0, maxDataColumnCount);
            }

            // Save the workbook
            workbook.Save("FreezeColumnsResult.xlsx");
        }
    }
}
