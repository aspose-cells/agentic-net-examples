// Title: Aspose.Cells C# Example – Move a Pivot Table Field from Column Area to Row Area
// Description: Demonstrates how to programmatically relocate a pivot field (e.g., "Region") from the column area to the row area using Aspose.Cells for .NET. The sample creates a workbook, adds sample data, builds a pivot table, removes the field with RemoveField, adds it back with AddFieldToArea, then refreshes and recalculates the pivot before saving the file.
// Keywords: Aspose.Cells | C# | PivotTable | move pivot field | column to row | RemoveField | AddFieldToArea | RefreshData | CalculateData | Excel automation | programmatic pivot layout | GitHub example | .NET Excel library
// Common Searches: Aspose.Cells move pivot field column to row | C# change pivot table field area programmatically | RemoveField AddFieldToArea Aspose.Cells example | how to refresh pivot after moving field Aspose.Cells | pivot table layout change Aspose.Cells .NET
// Developer Intent: Reposition a pivot table field from the column area to the row area using Aspose.Cells APIs.
// Use Cases: Allow end‑users to toggle pivot dimensions (e.g., switch Region between columns and rows) at runtime. | Generate reports where the layout must adapt to different visual preferences before exporting to Excel. | Automate workbook preparation for printing or dashboard publishing by programmatically adjusting pivot orientation.
// AI Prompts: Provide C# code that moves a pivot field from the column area to the row area with Aspose.Cells and updates the pivot data. | Explain the sequence of Aspose.Cells methods required to change a pivot field’s parent area and recalculate the table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemMoveDemo
{
    // Demonstrates how to programmatically relocate a pivot field (e.g., "Region") from the column area to the row area using Aspose.Cells for .NET. The sample creates a workbook, adds sample data, builds a pivot table, removes the field with RemoveField, adds it back with AddFieldToArea, then refreshes and recalculates the pivot before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Product, Region, Sales
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Apple");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(950);

            // Add a pivot table based on the data range
            // Destination top‑left cell is E3
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Initially place "Product" in the Row area and "Region" in the Column area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            // Add the data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate to populate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // Move the "Region" field from Column area to Row area
            // This changes the parent field of its pivot items.
            // ------------------------------------------------------------

            // Remove the field from the Column area
            pivotTable.RemoveField(PivotFieldType.Column, "Region");

            // Add the same field to the Row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");

            // Re‑calculate after the structural change
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotItemMoved.xlsx");
        }
    }
}
