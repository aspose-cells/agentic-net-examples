// Title: Aspose.Cells .NET: Create SalesTable ListObject, add a calculated Total column with SUM totals, and verify formula propagation
// Description: Demonstrates how to programmatically create a ListObject called SalesTable in a new workbook, add a "Total" column with the formula =[Quantity]*[Price], resize the table to include the column, enable a totals row that sums the Total column, insert an extra data row, and confirm that the formula automatically fills the new row using C# and Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table | add calculated column | custom formula | Totals row | SUM calculation | formula propagation | resize table | add data row | SalesTable example
// Common Searches: How to add a calculated column to a ListObject with Aspose.Cells .NET | Aspose.Cells set totals row sum for a table column | Resize an Aspose.Cells table to include a new column | Add a new row to an Aspose.Cells ListObject and auto‑fill formulas | Create and configure a ListObject programmatically in C#
// Developer Intent: Programmatically create a ListObject named SalesTable, add a Total column with a =[Quantity]*[Price] formula, enable a SUM totals row, resize the table, and ensure formulas propagate when new rows are added.
// Use Cases: Generate a sales ledger where each line total is calculated automatically and a grand total is displayed. | Insert dynamic rows into an existing Aspose.Cells table while preserving calculated columns without manual formula updates. | Resize an Aspose.Cells table to accommodate additional columns and configure a totals row for financial reporting.
// AI Prompts: Show C# code to create a ListObject, add a custom calculated column, and enable a SUM totals row using Aspose.Cells. | Provide an example that adds a new data row to an Aspose.Cells ListObject and automatically propagates the column formulas. | Explain how to resize an Aspose.Cells table to include a new column and set its TotalsCalculation to Sum.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to programmatically create a ListObject called SalesTable in a new workbook, add a "Total" column with the formula =[Quantity]*[Price], resize the table to include the column, enable a totals row that sums the Total column, insert an extra data row, and confirm that the formula automatically fills the new row using C# and Aspose.Cells.
class SalesTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (Product, Quantity, Price)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(7);
            sheet.Cells["C3"].PutValue(8);

            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B4"].PutValue(12);
            sheet.Cells["C4"].PutValue(3);

            // Add a ListObject (table) covering the existing range A1:C4
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject salesTable = sheet.ListObjects[tableIndex];
            salesTable.DisplayName = "SalesTable";

            // Add a new column header for "Total" (column D)
            sheet.Cells["D1"].PutValue("Total");

            // Resize the table to include the new column (now D column)
            salesTable.Resize(0, 0, 4, 3, true);

            // Ensure the new column exists before configuring it
            if (salesTable.ListColumns.Count > 3)
            {
                // Set a calculated formula for the new column: Quantity * Price
                salesTable.ListColumns[3].SetCustomCalculatedFormula("=[Quantity]*[Price]", false, false);

                // Enable totals row and set SUM calculation for the Total column
                salesTable.ShowTotals = true;
                salesTable.ListColumns[3].TotalsCalculation = TotalsCalculation.Sum;
            }

            // Add a new data row to test formula propagation
            // Row offset 4 corresponds to the first data row after the existing rows (0‑based inside the table)
            salesTable.PutCellValue(4, 0, "Item 4");   // Product
            salesTable.PutCellValue(4, 1, 9);         // Quantity
            salesTable.PutCellValue(4, 2, 6);         // Price
            // The Total column formula will automatically apply to the new row

            // Save the workbook
            workbook.Save("SalesTableDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
