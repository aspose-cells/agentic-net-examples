// Title: C# – Cut and paste a pivot table with Aspose.Cells while preserving its structure
// Description: Shows how to build a workbook, create a pivot table, and move the whole pivot table to a different cell using PivotTable.MoveTo in Aspose.Cells for .NET, keeping the pivot cache and formatting unchanged.
// Keywords: Aspose.Cells pivot table move | C# PivotTable.MoveTo | cut paste pivot table .NET | preserve pivot cache Aspose | relocate pivot table range | Aspose.Cells example C# | move pivot table to new location | Aspose.Cells range operations
// Common Searches: move pivot table Aspose.Cells C# | cut and paste pivot table preserving cache | PivotTable.MoveTo example .NET | relocate pivot table without rebuilding | Aspose.Cells cut range with pivot
// Developer Intent: The developer needs to cut a range that contains a pivot table and paste it elsewhere, keeping the pivot’s structure and cache intact.
// Use Cases: Re‑position a pivot table after inserting rows or columns in an automated report. | Place a pivot table on a different part of the sheet for better layout without recreating the cache. | Move a pivot table to another worksheet for separate presentation while retaining calculations.
// AI Prompts: Generate C# code using Aspose.Cells to cut a pivot table from cell D2 and paste it to B10, preserving the pivot cache. | Explain the parameters and behavior of PivotTable.MoveTo in Aspose.Cells for .NET. | Provide an example that moves multiple pivot tables in a workbook to new locations without losing data or formatting.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCutPasteDemo
{
    // Shows how to build a workbook, create a pivot table, and move the whole pivot table to a different cell using PivotTable.MoveTo in Aspose.Cells for .NET, keeping the pivot cache and formatting unchanged.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data that will be used as the pivot table source
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Orange");
            sheet.Cells["C3"].PutValue(850);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(560);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Tomato");
            sheet.Cells["C5"].PutValue(730);

            // ------------------------------------------------------------
            // Create a pivot table based on the data range A1:C5
            // The pivot table will be placed initially at D2
            // ------------------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Sales as data
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate data so the pivot table is populated
            pivot.CalculateData();

            // ------------------------------------------------------------
            // Move (cut & paste) the entire pivot table to a new location
            // Here we move it to row 10, column 2 (i.e., cell B10)
            // The MoveTo method preserves the pivot structure and cache
            // ------------------------------------------------------------
            pivot.MoveTo(9, 1); // Zero‑based indices: row 9 = 10th row, column 1 = B

            // After moving, recalculate to refresh the view at the new location
            pivot.CalculateData();

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("PivotTableCutPasteResult.xlsx");
        }
    }
}
