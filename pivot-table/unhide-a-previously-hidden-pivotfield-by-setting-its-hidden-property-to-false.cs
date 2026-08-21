// Title: Unhide a PivotField Item in Aspose.Cells for .NET Pivot Tables (HideItem false)
// Description: Demonstrates how to create a workbook, add a pivot table, hide a row‑field item with HideItem("Vegetable", true), then reveal the same item by calling HideItem("Vegetable", false) followed by RefreshData and CalculateData, and finally save the file.
// Keywords: Aspose.Cells | .NET | C# | PivotTable | HideItem | unhide pivot item | show hidden pivot field | RefreshData | CalculateData | programmatic pivot filter
// Common Searches: Aspose.Cells unhide pivot field item | HideItem false example C# | how to show hidden row field in Aspose pivot table | refresh pivot after unhiding item Aspose.Cells | C# code to toggle pivot item visibility
// Developer Intent: Reveal a previously hidden pivot field item by setting its hidden flag to false and updating the pivot table.
// Use Cases: Allow users to hide categories during data loading and automatically display them once processing completes. | Implement a toggle button that switches a pivot table row item on or off without rebuilding the table. | Reset a filtered report by programmatically unhiding all hidden pivot items before exporting.
// AI Prompts: Write C# code using Aspose.Cells to unhide the pivot field item "Vegetable" and explain why RefreshData and CalculateData are required. | Compare HideItem(string, true) and HideItem(string, false) in Aspose.Cells and describe their impact on pivot table rendering. | Provide a step‑by‑step guide to hide multiple pivot field items and later unhide them using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnhideDemo
{
    // Demonstrates how to create a workbook, add a pivot table, hide a row‑field item with HideItem("Vegetable", true), then reveal the same item by calling HideItem("Vegetable", false) followed by RefreshData and CalculateData, and finally save the file.
    class Program
    {
        static void Main()
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

            // Add the "Category" field to the row area and "Quantity" to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Get the row field (the one we will hide/unhide)
            PivotField rowField = pivotTable.RowFields[0];

            // -------------------------------------------------
            // Hide the item "Vegetable" using HideItem(string, true)
            // -------------------------------------------------
            rowField.HideItem("Vegetable", true);

            // Refresh and calculate to apply the hide operation
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Unhide the previously hidden item "Vegetable"
            // by setting its hidden flag to false via HideItem(string, false)
            // -------------------------------------------------
            rowField.HideItem("Vegetable", false);

            // Refresh and calculate again to reflect the unhide operation
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the updated pivot table
            workbook.Save("PivotField_UnhideDemo.xlsx");
        }
    }
}
