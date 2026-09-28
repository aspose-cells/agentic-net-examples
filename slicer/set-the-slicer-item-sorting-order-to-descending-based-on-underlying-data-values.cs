// Title: How to set a slicer’s items to sort descending in an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table, inserts a slicer linked to a row field, and sets the slicer’s SortOrderType to Descending with Aspose.Cells. | Write a .NET example that demonstrates configuring a slicer to sort its items in descending order after refreshing a pivot table using Aspose.Cells. | Provide a step‑by‑step C# snippet that adds a slicer to a worksheet, links it to a pivot table field, applies descending sort to the slicer items, and saves the file.
// Common Searches: Aspose.Cells C# how to sort slicer items descending in a pivot table | C# code sample for setting slicer sort order to descending with Aspose.Cells | descending sort order for Excel slicer using Aspose.Cells .NET API | set slicer SortOrderType to Descending in Aspose.Cells example
// Tags: Aspose.Cells slicer descending sort | C# pivot table slicer configuration | Set slicer SortOrderType Descending .NET | Excel slicer item sorting Aspose.Cells | Aspose.Cells workbook pivot slicer example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates creating a workbook, building a pivot table, adding a slicer linked to the row field, setting the slicer’s SortOrderType to Descending, and saving the workbook as SlicerSortedDesc.xlsx using Aspose.Cells for .NET.
class SetSlicerSortOrder
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(150);

            // Create a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table using the correct API
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a slicer linked to the first row field of the pivot table
            // Position the slicer at cell E1 (row 0, column 4 – zero‑based indexes)
            int slicerIndex = sheet.Slicers.Add(pivotTable, 0, 4, "CategorySlicer");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Set the slicer items to be sorted in descending order
            slicer.SortOrderType = SortOrder.Descending;

            // Save the workbook with the configured slicer
            workbook.Save("SlicerSortedDesc.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
