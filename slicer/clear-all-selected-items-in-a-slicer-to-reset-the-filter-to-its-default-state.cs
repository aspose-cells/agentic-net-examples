using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Fruit";
        cells["A2"].Value = "Apple";
        cells["A3"].Value = "Banana";
        cells["A4"].Value = "Apple";
        cells["A5"].Value = "Banana";

        cells["B1"].Value = "Sales";
        cells["B2"].Value = 120;
        cells["B3"].Value = 150;
        cells["B4"].Value = 130;
        cells["B5"].Value = 170;

        // Add a pivot table based on the data
        int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "FruitPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table
        SlicerCollection slicers = sheet.Slicers;
        int slicerIdx = slicers.Add(pivot, "F1", "Fruit");
        Slicer slicer = slicers[slicerIdx];

        // ---- Clear all selected items in the slicer ----
        // Iterate through each cache item and deselect it
        foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
        {
            item.Selected = false;
        }

        // Refresh the slicer so the pivot table reflects the change
        slicer.Refresh();

        // Save the workbook
        workbook.Save("ClearedSlicer.xlsx");
    }
}