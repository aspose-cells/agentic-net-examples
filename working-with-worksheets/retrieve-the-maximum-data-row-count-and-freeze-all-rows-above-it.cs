// Title: C# Aspose.Cells: Retrieve MaxDataRow and Freeze All Upper Rows
// Description: Shows how to call Cells.MaxDataRow to get the zero‑based index of the final data row, compute the first scrollable row, and use Worksheet.FreezePanes to lock every row above it before saving the workbook as an XLSX file.
// Keywords: Aspose.Cells MaxDataRow | Worksheet.FreezePanes C# | freeze top rows Excel .NET | dynamic header freeze Aspose | C# Excel row index
// Common Searches: Aspose.Cells get last data row C# | Freeze rows based on data range Aspose.Cells | Worksheet.FreezePanes example .NET | How to lock header rows dynamically in Excel | MaxDataRow property usage Aspose
// Developer Intent: Identify the last row containing data and apply a freeze pane that locks all rows above it.
// Use Cases: Create reports where headers stay visible while scrolling large tables. | Export variable‑size datasets with fixed top rows for better readability. | Implement pagination logic that requires knowledge of the final data row before freezing panes.
// AI Prompts: Generate C# code using Aspose.Cells to find the last populated row and freeze every row above it. | Explain the parameters of Worksheet.FreezePanes in relation to Cells.MaxDataRow. | Adapt the example to also freeze columns up to the last populated column.

using System;
using Aspose.Cells;

// Shows how to call Cells.MaxDataRow to get the zero‑based index of the final data row, compute the first scrollable row, and use Worksheet.FreezePanes to lock every row above it before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data (replace with your own data loading if needed)
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Item 1");
        cells["A3"].PutValue("Item 2");
        cells["A4"].PutValue("Item 3");

        // Retrieve the maximum data row index (zero‑based). Returns -1 if no data.
        int maxDataRow = cells.MaxDataRow;

        if (maxDataRow >= 0)
        {
            // Freeze all rows up to and including the max data row.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Split occurs at the first empty row after the data.
            int splitRow = maxDataRow + 1; // first row that should remain scrollable
            sheet.FreezePanes(splitRow, 0, splitRow, 0);
        }

        // Save the workbook
        workbook.Save("FrozenRows.xlsx", SaveFormat.Xlsx);
    }
}
