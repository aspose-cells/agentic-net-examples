// Title: How to rename a hidden PivotTable row field to a new label using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to hide all items of a PivotTable row field except a specific value and then change the field’s display name. | Programmatically update the name of a PivotField after applying HideItem filters and refresh the pivot table with Aspose.Cells.
// Common Searches: Aspose.Cells C# rename pivot row field after hiding items | change PivotField name after HideItem filter in .NET | how to update pivot table field label programmatically with Aspose.Cells | C# hide pivot items then set new field caption using Aspose.Cells
// Tags: pivotfield rename after hideitem aspnet | c# update pivot table field caption | excel pivotfield hide items programmatically | aspnet pivot table row field rename | aspose.cells pivot field label change

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRenameDemo
{
    // The example creates a workbook, adds product data, builds a PivotTable, hides all row items except "Widget", renames the row field to "Item Category", refreshes and calculates the pivot, and saves the file as PivotFieldRenamedAfterHide.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Widget";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "Gadget";
            cells["B3"].Value = 850;
            cells["A4"].Value = "Doohickey";
            cells["B4"].Value = 430;
            cells["A5"].Value = "Widget";
            cells["B5"].Value = 760;

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Access the row field
            PivotField productField = pivotTable.RowFields[0];

            // Hide all products except "Widget"
            for (int i = 0; i < productField.ItemCount; i++)
            {
                // Hide the item if its name is not "Widget"
                productField.HideItem(i, productField.Items[i] != "Widget");
            }

            // After hiding, rename the field to match new business terminology
            productField.Name = "Item Category";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook (lifecycle save)
            workbook.Save("PivotFieldRenamedAfterHide.xlsx");
        }
    }
}
