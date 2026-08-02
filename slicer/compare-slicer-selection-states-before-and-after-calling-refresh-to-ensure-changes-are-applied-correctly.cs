using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerRefreshComparison
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table's "Category" field
        int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Category");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Initial selection: select only the first slicer item
        for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
        {
            slicer.SlicerCache.SlicerCacheItems[i].Selected = (i == 0);
        }

        // Capture selection states before refresh
        List<bool> beforeRefresh = slicer.SlicerCache.SlicerCacheItems
            .Select(item => item.Selected).ToList();

        Console.WriteLine("Selection states BEFORE Refresh:");
        PrintSelectionStates(slicer);

        // Modify the source data by adding a new category
        sheet.Cells["A5"].PutValue("D");
        sheet.Cells["B5"].PutValue(40);

        // Refresh the slicer (also refreshes the underlying pivot table)
        slicer.Refresh();

        // Capture selection states after refresh
        List<bool> afterRefresh = slicer.SlicerCache.SlicerCacheItems
            .Select(item => item.Selected).ToList();

        Console.WriteLine("\nSelection states AFTER Refresh:");
        PrintSelectionStates(slicer);

        // Compare the two states
        bool statesUnchanged = beforeRefresh.Count == afterRefresh.Count &&
                               !beforeRefresh.Where((val, idx) => val != afterRefresh[idx]).Any();

        Console.WriteLine($"\nSelection states unchanged: {statesUnchanged}");

        // Save the workbook
        workbook.Save("SlicerRefreshComparison.xlsx");
    }

    // Helper method to print each slicer item's value and selection flag
    static void PrintSelectionStates(Slicer slicer)
    {
        for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
        {
            SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
            Console.WriteLine($"Item '{item.Value}': Selected = {item.Selected}");
        }
    }
}