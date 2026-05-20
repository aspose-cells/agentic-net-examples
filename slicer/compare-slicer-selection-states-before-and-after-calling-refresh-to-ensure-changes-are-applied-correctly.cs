using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRefreshComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
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
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Ensure initial selection: select first item, deselect others
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                slicer.SlicerCache.SlicerCacheItems[i].Selected = i == 0;
            }

            // Capture selection states before refresh
            List<bool> beforeRefresh = new List<bool>();
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                beforeRefresh.Add(item.Selected);
            }

            // Modify underlying data: add a new category that will appear in the slicer
            dataSheet.Cells["A5"].PutValue("D");
            dataSheet.Cells["B5"].PutValue(40);

            // Refresh the slicer (also refreshes the pivot table)
            slicer.Refresh();

            // Capture selection states after refresh
            List<bool> afterRefresh = new List<bool>();
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                afterRefresh.Add(item.Selected);
            }

            // Compare and output differences
            Console.WriteLine("Comparison of slicer selection states before and after Refresh:");
            int count = Math.Max(beforeRefresh.Count, afterRefresh.Count);
            for (int i = 0; i < count; i++)
            {
                string itemValue = i < slicer.SlicerCache.SlicerCacheItems.Count
                    ? slicer.SlicerCache.SlicerCacheItems[i].Value
                    : $"Item{i}";
                bool before = i < beforeRefresh.Count ? beforeRefresh[i] : false;
                bool after = i < afterRefresh.Count ? afterRefresh[i] : false;
                Console.WriteLine($"Item '{itemValue}': before={before}, after={after}");
            }

            // Save the workbook
            workbook.Save("SlicerRefreshComparison.xlsx");
        }
    }
}