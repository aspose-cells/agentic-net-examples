// Title: Create an Excel table with structured reference formulas and a calculated column using Aspose.Cells for .NET (C#)
// Common Searches: how to use structured references in Aspose.Cells to sum a column in a ListObject total row C# | adding a calculated column with row-level structured reference in an Aspose.Cells table | resize an Aspose.Cells ListObject to include a new column while keeping formulas dynamic | Aspose.Cells C# example of SUM formula using table name and column header | populate total price column using [@Quantity]*[@Price] in Aspose.Cells
// Tags: Aspose.Cells structured reference SUM formula | C# ListObject calculated column | Aspose.Cells table resize add column | Excel table total row formula Aspose.Cells | row-level structured reference formula C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject

// The example creates a workbook, defines a ListObject named SalesData over product data, enables the total row, and inserts a SUM formula that uses a structured reference (SalesData[Quantity]). It then adds a new column "TotalPrice", resizes the table to include it, and fills each cell with a calculated formula using row-level structured references ([@Quantity]*[@Price]). Finally, the workbook is saved as StructuredReferenceExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // ----- Populate sample data -----
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue(0.8);

            // ----- Create a table (ListObject) over the data range A1:C4 -----
            // Parameters are zero‑based indices: firstRow, firstColumn, lastRow, lastColumn, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set table name (DisplayName works across Aspose.Cells versions)
            table.DisplayName = "SalesData";

            // Enable the total row
            table.ShowTotals = true;

            // ----- Use structured reference in the total row -----
            // Quantity column is the second column (index 1) within the table
            int quantityColumnIndex = 1;
            // Total row is the last row of the table when ShowTotals is true
            int totalRowIndex = table.EndRow;
            sheet.Cells[totalRowIndex, table.StartColumn + quantityColumnIndex].Formula =
                $"=SUM({table.DisplayName}[Quantity])";

            // ----- Add a calculated column that uses structured references -----
            // Insert a new column header for TotalPrice (column D)
            int totalPriceColumnIndex = 3; // zero‑based column index for D
            sheet.Cells[0, totalPriceColumnIndex].PutValue("TotalPrice");

            // Expand the table to include the new column
            int currentRows = table.EndRow - table.StartRow + 1; // includes header, data, and total row
            int newColumns = table.EndColumn - table.StartColumn + 2; // add one column
            // The last parameter indicates that the table has headers
            table.Resize(table.StartRow, table.StartColumn, currentRows, newColumns, true);

            // Apply a formula to each cell in the new column: =[@Quantity]*[@Price]
            int newColumnIndex = totalPriceColumnIndex; // after resizing this is the new column
            int firstDataRow = table.StartRow + 1; // row after header
            int lastDataRow = table.EndRow - (table.ShowTotals ? 1 : 0); // exclude total row
            for (int r = firstDataRow; r <= lastDataRow; r++)
            {
                sheet.Cells[r, newColumnIndex].Formula = "=[@Quantity]*[@Price]";
            }

            // ----- Save the workbook -----
            string outputPath = "StructuredReferenceExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
