// Title: How to apply a custom descending alphabetical sort to a row field in an Aspose.Cells pivot table using C#
// AI Prompts: Write C# code that creates a workbook, adds a pivot table from a range, and uses PivotField.SortBy to sort the row field labels from Z to A with Aspose.Cells. | Show the steps to refresh and recalculate a pivot table after applying a descending alphabetical sort to its row items in .NET. | Demonstrate how to save the workbook after sorting pivot table categories in descending order using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to sort pivot table row items alphabetically descending | programmatically set Z‑to‑A sort on pivot table category field using Aspose.Cells | C# example for custom descending sort of pivot table row field labels Aspose.Cells
// Tags: pivot field SortBy descending Aspose.Cells | C# Aspose.Cells row field alphabetical ordering | custom pivot table sorting .NET | Excel workbook save after pivot sort Aspose | pivot table refresh post sort Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    // The program creates a workbook, fills it with sample category‑sales data, adds a pivot table, places the Category field in the row area, applies a descending alphabetical sort to the Category pivot items using PivotField.SortBy, refreshes and calculates the pivot table, and finally saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with product categories
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = 1200;
            sheet.Cells["A3"].Value = "Furniture";
            sheet.Cells["B3"].Value = 800;
            sheet.Cells["A4"].Value = "Clothing";
            sheet.Cells["B4"].Value = 450;
            sheet.Cells["A5"].Value = "Books";
            sheet.Cells["B5"].Value = 300;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIdx];

            // Add the Category field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply a custom descending alphabetical sort to the Category pivot items
            // -1 indicates sorting by the field's own labels (i.e., alphabetical order)
            // SortOrder.Descending makes it descending (Z → A)
            PivotField categoryField = pivotTable.RowFields[0];
            categoryField.SortBy(SortOrder.Descending, -1);

            // Refresh and calculate the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomDescendingAlphabeticalSort.xlsx");
        }
    }
}
