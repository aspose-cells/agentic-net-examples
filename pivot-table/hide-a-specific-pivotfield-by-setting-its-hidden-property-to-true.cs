// Title: Hide a PivotField in an Aspose.Cells PivotTable (C#) – Set Items to Hidden
// Description: Demonstrates how to hide a specific PivotField in an Aspose.Cells PivotTable by iterating its items and calling HideItem(index, true). The example creates a workbook, adds sample data, builds a pivot table, hides the row field "Category", refreshes the table, and saves the result.
// Keywords: Aspose.Cells hide pivot field C# | PivotTable HideItem method | set pivot field hidden Aspose | C# hide row field in pivot table | Aspose.Cells PivotField visibility
// Common Searches: how to hide a pivot field in Aspose.Cells C# | Aspose.Cells hide row items in pivot table | C# hide specific field in Aspose pivot table | Aspose.Cells PivotTable hide field programmatically
// Developer Intent: Programmatically hide a specific PivotField so it does not appear in the generated PivotTable.
// Use Cases: Exclude confidential categories from a sales summary while keeping totals. | Create a clean report view by removing unnecessary row fields on the fly. | Allow end‑users to toggle visibility of pivot fields based on preferences in a .NET dashboard.
// AI Prompts: Generate C# code using Aspose.Cells to hide a column PivotField by setting each item's hidden flag. | Show how to toggle visibility of multiple PivotFields in an Aspose.Cells workbook at runtime. | Provide an example that hides a PivotField and then reveals it again using Aspose.Cells API.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideFieldDemo
{
    // Demonstrates how to hide a specific PivotField in an Aspose.Cells PivotTable by iterating its items and calling HideItem(index, true). The example creates a workbook, adds sample data, builds a pivot table, hides the row field "Category", refreshes the table, and saves the result.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["A5"].PutValue("Vegetable");

            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["B5"].PutValue(25);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Quantity" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Get the pivot field that we want to hide (the row field "Category")
            PivotField fieldToHide = pivotTable.RowFields[0];

            // Hide all items of this field by setting each item's hidden flag to true
            // This effectively hides the entire field from the pivot view
            for (int i = 0; i < fieldToHide.ItemCount; i++)
            {
                // Use the HideItem method (int index, bool isHidden) as defined in the API
                fieldToHide.HideItem(i, true);
            }

            // Refresh and calculate the pivot table after modifying visibility
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the hidden field
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}
