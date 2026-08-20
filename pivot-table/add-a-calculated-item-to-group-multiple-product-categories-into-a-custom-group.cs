// Title: Create a Calculated Item to Group Products in an Aspose.Cells Pivot Table (C#)
// Description: Shows how to generate a workbook with product‑sales data, add a pivot table, and define a calculated item that merges Apple and Banana into a custom row group called “FruitGroup”. The pivot is refreshed, recalculated, and saved as PivotCalculatedGroup.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# pivot table | calculated item | custom row group | Aspose.Cells for .NET | pivot table grouping | Excel pivot calculated item | PivotTable.AddCalculatedItem | Aspose.Cells API
// Common Searches: Aspose.Cells add calculated item C# | group rows in pivot table using Aspose.Cells | custom row group pivot Aspose.Cells .NET | how to sum specific items in Aspose.Cells pivot | create product group in Excel pivot with Aspose.Cells
// Developer Intent: Add a calculated item that aggregates selected row entries into a single group within an Aspose.Cells pivot table using C#.
// Use Cases: Summarize Apple and Banana sales together as “FruitGroup” for concise reporting. | Define additional groups (e.g., CitrusGroup) by adding more calculated items. | Refresh and recalculate the pivot after adding items to ensure accurate totals. | Automate custom grouping for dynamic dashboards and exported Excel files.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table and add a calculated item that groups Apple and Banana as FruitGroup. | Explain the steps to refresh and recalculate a pivot table after inserting a calculated item in Aspose.Cells for .NET. | Show how to add multiple calculated items for different product categories in a single Aspose.Cells pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to generate a workbook with product‑sales data, add a pivot table, and define a calculated item that merges Apple and Banana into a custom row group called “FruitGroup”. The pivot is refreshed, recalculated, and saved as PivotCalculatedGroup.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: Product names and their sales values
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["B2"].Value = 1000;

        sheet.Cells["A3"].Value = "Banana";
        sheet.Cells["B3"].Value = 1500;

        sheet.Cells["A4"].Value = "Orange";
        sheet.Cells["B4"].Value = 2000;

        sheet.Cells["A5"].Value = "Grapes";
        sheet.Cells["B5"].Value = 2500;

        // Create a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Product field to the row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

        // Add the Sales field to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field (Product) to which we will add a calculated item
        PivotField productField = pivotTable.RowFields[0];

        // Add a calculated item that groups Apple and Banana into a custom group named "FruitGroup"
        // The formula references the existing pivot items by their names
        productField.AddCalculatedItem("FruitGroup", "=Apple + Banana");

        // Refresh the pivot table data and recalculate to reflect the new calculated item
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the pivot table and custom calculated group
        workbook.Save("PivotCalculatedGroup.xlsx");
    }
}
