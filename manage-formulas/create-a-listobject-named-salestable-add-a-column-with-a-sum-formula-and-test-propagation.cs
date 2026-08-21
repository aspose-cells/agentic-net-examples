// Title: Aspose.Cells for .NET – Create a ListObject (SalesTable) with a calculated column, totals row, and automatic formula propagation (C#)
// Description: Demonstrates how to build a workbook, define a ListObject named SalesTable, add a totals row that sums the Sales column, set a column formula (DoubleSales = Sales × 2), insert a new row, and let Aspose.Cells automatically propagate the formula before saving the file.
// Keywords: Aspose.Cells ListObject C# | Create table Aspose.Cells | calculated column formula Aspose.Cells | totals row sum Aspose.Cells | formula propagation ListObject | C# Excel table example | Aspose.Cells add row auto‑fill formula | SalesTable Aspose.Cells
// Common Searches: Aspose.Cells add totals row to ListObject | Set column formula in Aspose.Cells table | Auto‑expand ListObject with formula C# | Create named table Aspose.Cells .NET | How to sum column in Aspose.Cells table
// Developer Intent: The developer wants to create a named ListObject, add a calculated column and a totals row that sums a data column, and ensure the formula automatically applies to newly added rows.
// Use Cases: Generate a sales worksheet where each entry shows double the original sales amount and the table footer provides the total sales. | Maintain calculated columns in dynamic reports without manually copying formulas when new product rows are added. | Export a structured Excel file with a named table and built‑in aggregation for quick data analysis.
// AI Prompts: Write C# code using Aspose.Cells to create a ListObject called 'SalesTable' with a calculated column that multiplies the 'Sales' field by 2 and a totals row that sums the 'Sales' column. | Explain how Aspose.Cells propagates a column formula in a ListObject when a new row is inserted and how to verify the result programmatically. | Provide troubleshooting steps if the formula does not auto‑fill after adding rows to the ListObject in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to build a workbook, define a ListObject named SalesTable, add a totals row that sums the Sales column, set a column formula (DoubleSales = Sales × 2), insert a new row, and let Aspose.Cells automatically propagate the formula before saving the file.
class SalesTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with headers
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B4"].PutValue(200);

            // Write the header for the calculated column
            sheet.Cells["C1"].PutValue("DoubleSales");

            // Create a ListObject (table) that covers the data range including the new column
            // startRow=0, startColumn=0, endRow=4 (last data row), endColumn=2 (includes column C), hasHeaders=true
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject salesTable = sheet.ListObjects[tableIndex];

            // Give the table a name
            salesTable.DisplayName = "SalesTable";

            // Show the totals row
            salesTable.ShowTotals = true;

            // Set the totals calculation for the "Sales" column (index 1) to Sum
            salesTable.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

            // Set the formula for the calculated column (index 2) so each cell equals Sales*2
            salesTable.ListColumns[2].Formula = "=[Sales]*2";

            // Add a new data row; the table will expand automatically and the formula will propagate
            salesTable.PutCellValue(3, 0, "Item 4");   // Product
            salesTable.PutCellValue(3, 1, 250);       // Sales
            // DoubleSales will be calculated as 250*2 = 500

            // Save the workbook
            string outputPath = "SalesTableDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
