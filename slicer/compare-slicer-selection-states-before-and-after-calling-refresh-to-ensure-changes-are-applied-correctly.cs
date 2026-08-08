// Title: Aspose.Cells .NET: Compare Slicer Selection Before and After Refresh
// Description: Shows how to record slicer item selections, change a selection, call Slicer.Refresh, and then compare the before‑and‑after states in a workbook with a pivot table using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# slicer refresh | slicer selection state | pivot table slicer | SlicerCacheItem | compare slicer before after | Excel slicer automation | Aspose.Cells API
// Common Searches: Aspose.Cells compare slicer selection before after refresh | how to check slicer state after Refresh in .NET | retrieve slicer cache items Aspose.Cells C# | preserve slicer selections when refreshing pivot table | debug slicer filtering with Aspose.Cells
// Developer Intent: Confirm that programmatic changes to slicer selections are reflected correctly after calling Refresh.
// Use Cases: Automated test to ensure slicer selections survive a pivot table refresh. | Logging differences in slicer item states for debugging filter logic. | Generating a report of selection changes when updating slicer values programmatically.
// AI Prompts: Write C# code using Aspose.Cells that captures slicer selections, toggles an item, refreshes the slicer, and prints the before‑after comparison. | Create a method that returns a list of slicer items whose Selected property changed after calling Slicer.Refresh in Aspose.Cells for .NET. | Explain how to maintain slicer selections when refreshing a linked pivot table and how to compare the selection states with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRefreshComparisonDemo
{
    // Shows how to record slicer item selections, change a selection, call Slicer.Refresh, and then compare the before‑and‑after states in a workbook with a pivot table using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the data
            int pivotIdx = dataSheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = dataSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = dataSheet.Slicers.Add(pivot, "F1", "Category");
            Slicer slicer = dataSheet.Slicers[slicerIdx];

            // Ensure all items are initially selected
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = true;
            }

            // Capture selection states before refresh
            bool[] beforeRefresh = new bool[slicer.SlicerCache.SlicerCacheItems.Count];
            for (int i = 0; i < beforeRefresh.Length; i++)
            {
                beforeRefresh[i] = slicer.SlicerCache.SlicerCacheItems[i].Selected;
            }

            // Change selection: deselect the first item, keep others selected
            if (slicer.SlicerCache.SlicerCacheItems.Count > 0)
            {
                slicer.SlicerCache.SlicerCacheItems[0].Selected = false;
            }

            // Refresh the slicer (also refreshes the underlying pivot table)
            slicer.Refresh();

            // Capture selection states after refresh
            bool[] afterRefresh = new bool[slicer.SlicerCache.SlicerCacheItems.Count];
            for (int i = 0; i < afterRefresh.Length; i++)
            {
                afterRefresh[i] = slicer.SlicerCache.SlicerCacheItems[i].Selected;
            }

            // Compare and output the differences
            Console.WriteLine("Comparison of slicer selection states (Before -> After):");
            for (int i = 0; i < beforeRefresh.Length; i++)
            {
                Console.WriteLine($"Item {i} ('{slicer.SlicerCache.SlicerCacheItems[i].Value}'): {beforeRefresh[i]} -> {afterRefresh[i]}");
            }

            // Save the workbook (required by lifecycle rule)
            workbook.Save("SlicerRefreshComparisonDemo.xlsx");
        }
    }
}
