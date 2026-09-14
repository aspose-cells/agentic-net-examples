// Title: Create a structured ListObject table from a 2‑D array and freeze its header row using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that populates a worksheet from a two‑dimensional string array and converts the range into a ListObject table. | Add a built‑in style to the ListObject, assign a custom display name, and save the workbook as an .xlsx file. | Configure the worksheet to freeze the table’s header row so it stays visible while scrolling.
// Common Searches: how to add a ListObject table from an array with Aspose.Cells C# | freeze first row of an Excel sheet using Aspose.Cells .NET | apply built‑in table style to a structured table in Aspose.Cells | set display name for ListObject table Aspose.Cells example | save workbook with frozen panes Aspose.Cells C#
// Tags: Aspose.Cells ListObject creation from 2D array | freeze panes for table header Aspose.Cells | assign built‑in style to ListObject Aspose.Cells | set ListObject display name Aspose.Cells | save workbook with frozen panes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// // This program creates a new workbook, writes a 2‑dimensional string array to the first worksheet, defines the range as a ListObject table named SalesTable, applies a medium built‑in style, freezes the header row by setting freeze panes at row 1, and saves the file as StructuredTableWithFrozenHeader.xlsx.
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
            sheet.Name = "Data";

            // Sample data including header row
            string[,] data = new string[,]
            {
                { "ID", "Name", "Quantity", "Price" },
                { "1", "Apple", "50", "0.5" },
                { "2", "Banana", "30", "0.3" },
                { "3", "Orange", "20", "0.6" },
                { "4", "Grape", "40", "0.8" }
            };

            // Insert data into the worksheet starting at cell A1
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    sheet.Cells[i, j].PutValue(data[i, j]);
                }
            }

            // Define the range that will become the structured table
            int firstRow = 0;
            int firstCol = 0;
            int totalRows = data.GetLength(0);
            int totalCols = data.GetLength(1);
            CellArea tableArea = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstCol,
                EndRow = firstRow + totalRows - 1,
                EndColumn = firstCol + totalCols - 1
            };

            // Add a ListObject (structured table) to the defined range
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set display name for the table (Name property may not be available in older versions)
            table.DisplayName = "SalesTable";

            // Show header row (default is true) and hide total row
            table.ShowHeaderRow = true;
            // Total row is hidden by default; no need to set ShowTotalRow if unavailable

            // Apply a built‑in style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Freeze the header row so it remains visible while scrolling
            // Freeze panes at row index 1 (second row) and column index 0 (first column)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("StructuredTableWithFrozenHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
