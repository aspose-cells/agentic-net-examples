// Title: Freeze columns up to the last used column using MaxDataColumn in Aspose.Cells for .NET
// Description: Demonstrates how to obtain the index of the final populated column with Worksheet.Cells.MaxDataColumn, apply Worksheet.FreezePanes to lock all columns to its left, and save the result as an Excel file. Ideal for creating scroll‑locked reports in C# worldwide.
// Keywords: Aspose.Cells | C# | MaxDataColumn | FreezePanes | freeze columns | last used column | Excel automation | worksheet freezing | dynamic column range
// Common Searches: Aspose.Cells freeze columns up to last data column | C# get last populated column Excel | FreezePanes based on MaxDataColumn example | How to lock columns in Aspose.Cells | Retrieve maximum data column index Aspose
// Developer Intent: Identify the final data column in a worksheet and freeze every column before it.
// Use Cases: Generate reports where header columns stay visible while scrolling horizontally. | Create templates with a variable number of data columns that are automatically locked for navigation. | Export large data sets and prevent users from scrolling beyond the populated area.
// AI Prompts: Write C# code with Aspose.Cells that finds the last column containing data and freezes all preceding columns. | Explain the relationship between MaxDataColumn and FreezePanes for column locking in an Excel workbook. | Provide a complete example that saves an Excel file after freezing columns based on the maximum data column.

using System;
using Aspose.Cells;

// Demonstrates how to obtain the index of the final populated column with Worksheet.Cells.MaxDataColumn, apply Worksheet.FreezePanes to lock all columns to its left, and save the result as an Excel file. Ideal for creating scroll‑locked reports in C# worldwide.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add some sample data to demonstrate the logic
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue(100);
        worksheet.Cells["B2"].PutValue(200);
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["C2"].PutValue(300);

        // Retrieve the maximum data column index (zero‑based)
        int maxDataColumn = worksheet.Cells.MaxDataColumn; // e.g., 2 for columns A‑C

        // Freeze all columns to the left of the first empty column.
        // FreezePanes(row, column, freezedRows, freezedColumns)
        // - row = 0 (no frozen rows)
        // - column = maxDataColumn + 1 (first column after the data)
        // - freezedRows = 0
        // - freezedColumns = maxDataColumn + 1 (number of columns to freeze)
        worksheet.FreezePanes(0, maxDataColumn + 1, 0, maxDataColumn + 1);

        // Save the workbook
        workbook.Save("FreezeColumns.xlsx");
    }
}
