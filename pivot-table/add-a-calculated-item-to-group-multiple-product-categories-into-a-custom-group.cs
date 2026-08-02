// Title: Create a Custom Calculated Item to Group Product Categories in an Aspose.Cells Pivot Table (C#)
// Description: This example builds a workbook with product and sales data, adds a pivot table, places the Product field in the row area, and defines a calculated item named "FruitGroup" that sums the Apple and Banana entries. The pivot is refreshed, calculated, and saved as an Excel file containing the custom group.
// Keywords: Aspose.Cells | C# pivot table | calculated item | custom group | group row items | pivot table formula | Excel aggregation | product category grouping
// Common Searches: Aspose.Cells add calculated item C# | group rows in pivot table Aspose.Cells | custom category group pivot Aspose | sum specific items in Aspose.Cells pivot | create FruitGroup calculated item
// Developer Intent: Add a calculated item that combines selected product rows into a single custom group within an Aspose.Cells pivot table using C#.
// Use Cases: Summarize sales of Apple and Banana under one "FruitGroup" row for concise reporting. | Create dynamic category groups without modifying the source data, enabling flexible dashboards. | Perform comparative analysis by aggregating chosen items while keeping other pivot fields intact.
// AI Prompts: Generate C# code with Aspose.Cells that adds a calculated item "FruitGroup" combining "Apple" and "Banana" in a pivot table. | Explain how calculated item formulas reference existing pivot items in Aspose.Cells. | Outline the steps to refresh and recalculate a pivot table after inserting a calculated item using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomGroupDemo
{
    // This example builds a workbook with product and sales data, adds a pivot table, places the Product field in the row area, and defines a calculated item named "FruitGroup" that sums the Apple and Banana entries. The pivot is refreshed, calculated, and saved as an Excel file containing the custom group.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product categories and their sales
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 1200;

            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = 800;

            sheet.Cells["A4"].Value = "Orange";
            sheet.Cells["B4"].Value = 1500;

            sheet.Cells["A5"].Value = "Grapes";
            sheet.Cells["B5"].Value = 900;

            sheet.Cells["A6"].Value = "Mango";
            sheet.Cells["B6"].Value = 1100;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "ProductPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Product field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row pivot field (Product)
            PivotField productField = pivotTable.RowFields[0];

            // Add a calculated item that groups Apple and Banana into a custom group named "FruitGroup"
            // The formula references the existing items by their names.
            productField.AddCalculatedItem("FruitGroup", "=Apple + Banana");

            // Refresh and calculate the pivot table to apply the new calculated item
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the pivot table and custom calculated group
            workbook.Save("ProductPivotWithCustomGroup.xlsx");
        }
    }
}
