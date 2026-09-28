// Title: How to enable a totals row and add a COUNT(DISTINCT) formula to an Aspose.Cells ListObject table in C#
// AI Prompts: Generate C# code that creates a worksheet, adds a ListObject, turns on its totals row, and inserts a COUNT(DISTINCT) formula for a specific column using Aspose.Cells. | Write a C# snippet with Aspose.Cells that displays a totals row in a table and sets a distinct‑count aggregation formula for the Category field.
// Common Searches: Aspose.Cells C# show totals row for ListObject table | C# Aspose.Cells count distinct values in table totals row | how to set custom formula in totals row of Aspose.Cells table | add totals row with distinct count aggregation using Aspose.Cells for .NET
// Tags: Aspose.Cells ListObject show totals row | Aspose.Cells set totals row formula | COUNT DISTINCT formula in Aspose.Cells table | C# Aspose.Cells table totals aggregation | Excel table distinct count using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, adds sample data, defines a ListObject table, enables its totals row, places a COUNT(DISTINCT([Category])) formula in the totals row cell, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (header + values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["A5"].PutValue("C");

            // Define the range for the table (including header row)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;         // rows 0‑4 (A1:B5)
            int totalColumns = 2;      // columns A and B

            // Add a ListObject (table) to the worksheet
            int tableIdx = sheet.ListObjects.Add(firstRow, firstColumn, firstRow + totalRows, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Set a display name for the table
            table.DisplayName = "MyTable";

            // Enable the totals row
            table.ShowTotals = true;

            // Set a custom formula in the totals row for the first column to count distinct values
            int totalsRowIndex = firstRow + totalRows; // row after the data rows
            sheet.Cells[totalsRowIndex, firstColumn].Formula = "COUNT(DISTINCT([Category]))";

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
