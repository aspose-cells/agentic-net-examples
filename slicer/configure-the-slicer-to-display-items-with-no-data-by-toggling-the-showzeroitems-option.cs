using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ConfigureSlicerShowZeroItems
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(0);   // Zero value (no data)
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("D");
            cells["B5"].PutValue(0);   // Zero value (no data)

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field
            // Note: The correct parameter order is (pivot, destination cell, field name)
            int slicerIdx = sheet.Slicers.Add(pivot, "E10", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Ensure the slicer shows all items, even those with no data
            slicer.ShowAllItems = true;

            // Configure how items with no data are displayed (e.g., natural order)
            slicer.ShowTypeOfItemsWithNoData = ItemsWithNoDataShowMode.Natural;

            // Optional: set a caption for clarity
            slicer.Caption = "Category Slicer (Show Zero Items)";

            // Refresh the slicer to apply changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerShowZeroItems.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}