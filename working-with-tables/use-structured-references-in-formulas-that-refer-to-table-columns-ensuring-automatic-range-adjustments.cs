// Title: Aspose.Cells .NET – Create Excel Table with Structured Reference Formula that Auto‑Expands
// Description: A C# example that builds a workbook, converts a data range into a ListObject named "Sales", adds a calculated "Total" column using the structured reference =[@Quantity]*[@Price], inserts an extra row, resizes the table to include the new data, recalculates all formulas, and saves the file as StructuredReferenceDemo.xlsx.
// Keywords: Aspose.Cells | C# Excel table | ListObject | structured reference formula | calculated column | auto expand table range | Resize ListObject | formula recalculation | Excel table automation | Aspose.Cells .NET
// Common Searches: structured reference formula Aspose.Cells | add calculated column to ListObject C# | auto expand Excel table after inserting rows Aspose | resize ListObject programmatically | Aspose.Cells calculate formulas | C# create Excel table with ListObject
// Developer Intent: Create an Excel ListObject, define a calculated column using a structured reference, and have the formula propagate automatically when new rows are added.
// Use Cases: Generate a “Total” column that multiplies Quantity by Price for each table row via a structured reference. | Append new sales records and let the Total column compute automatically without manual copying. | Programmatically resize the ListObject after inserting rows and trigger a full formula recalculation before saving.
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject, adds a calculated column with the formula =[@Quantity]*[@Price], and saves the workbook. | Show how to add a new row to an existing Aspose.Cells table and automatically extend the structured reference formula to the new row. | Explain the steps to resize a ListObject after adding data and force formula recalculation using Aspose.Cells .NET.

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// A C# example that builds a workbook, converts a data range into a ListObject named "Sales", adds a calculated "Total" column using the structured reference =[@Quantity]*[@Price], inserts an extra row, resizes the table to include the new data, recalculates all formulas, and saves the file as StructuredReferenceDemo.xlsx.
class StructuredReferenceDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ----- Populate sample data with headers -----
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            string[] items = { "Apple", "Banana", "Cherry" };
            int[] quantities = { 10, 20, 15 };
            double[] prices = { 0.5, 0.3, 0.8 };

            for (int i = 0; i < items.Length; i++)
            {
                cells[i + 1, 0].PutValue(items[i]);      // Column A
                cells[i + 1, 1].PutValue(quantities[i]); // Column B
                cells[i + 1, 2].PutValue(prices[i]);     // Column C
            }

            // ----- Create a table (ListObject) covering the data range -----
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = ws.ListObjects.Add(0, 0, items.Length, 2, true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "Sales";

            // ----- Add a new column "Total" to the table -----
            // Expand the table by one column (hasHeaders = true because the table already has a header row).
            table.Resize(table.StartRow, table.StartColumn, table.EndRow, table.EndColumn + 1, true);
            cells[0, 3].PutValue("Total"); // Header for the new column.

            // Set a column formula using a structured reference.
            // Structured reference syntax: =[@Quantity]*[@Price]
            int totalColumnIndex = table.ListColumns.Count - 1;
            table.ListColumns[totalColumnIndex].Formula = "=[@Quantity]*[@Price]";

            // ----- Add a new data row to demonstrate automatic formula propagation -----
            int newRowIndex = table.EndRow + 1; // Row index just below the current table.
            cells[newRowIndex, 0].PutValue("Date");
            cells[newRowIndex, 1].PutValue(5);
            cells[newRowIndex, 2].PutValue(1.2);
            // Expand the table to include the new row (hasHeaders = true).
            table.Resize(table.StartRow, table.StartColumn, newRowIndex, table.EndColumn, true);

            // Calculate all formulas so that the "Total" column values are updated.
            wb.CalculateFormula();

            // Save the workbook.
            string outputPath = "StructuredReferenceDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
