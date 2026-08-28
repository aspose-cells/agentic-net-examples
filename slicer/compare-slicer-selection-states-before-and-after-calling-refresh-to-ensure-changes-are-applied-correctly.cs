// Title: C# example: capture slicer selection flags, modify them, refresh the slicer, and compare before/after states with Aspose.Cells
// AI Prompts: Write C# code that iterates through a slicer's SlicerCacheItems, stores each item's Selected property, changes the selection to a specific item, calls slicer.Refresh, then prints which items changed. | Generate an Aspose.Cells .NET snippet that creates a pivot table, adds a linked slicer, toggles slicer selections, refreshes the slicer, and compares the selection flags before and after the refresh.
// Common Searches: Aspose.Cells how to get slicer selected values before refreshing the pivot table | C# compare slicer cache item selection state before and after slicer.Refresh | track changes in Aspose.Cells slicer selections after modifying cache items | example code for capturing slicer selection flags in Aspose.Cells .NET
// Tags: Aspose.Cells slicer cache state capture | C# refresh slicer linked to pivot | track slicer selection changes Aspose.Cells | Aspose.Cells slicer selection flag comparison | slicer.Refresh usage example

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRefreshComparisonDemo
{
    // The program creates a workbook, adds sample data, builds a pivot table, and attaches a slicer. It records the Selected flag of each slicer cache item, changes the selection to only the last item, calls slicer.Refresh (which also refreshes the pivot), records the flags again, prints a before/after comparison for each item, and saves the workbook as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["A2"].Value = "A";
            cells["A3"].Value = "B";
            cells["A4"].Value = "C";
            cells["B1"].Value = "Value";
            cells["B2"].Value = 10;
            cells["B3"].Value = 20;
            cells["B4"].Value = 30;

            // Create a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // -------------------------------------------------
            // Capture selection states BEFORE Refresh
            // -------------------------------------------------
            List<bool> beforeSelection = new List<bool>();
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                beforeSelection.Add(item.Selected);
            }

            // Change selection: select the last item, deselect others
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                slicer.SlicerCache.SlicerCacheItems[i].Selected = (i == slicer.SlicerCache.SlicerCacheItems.Count - 1);
            }

            // Refresh the slicer (also refreshes the underlying pivot table)
            slicer.Refresh();

            // -------------------------------------------------
            // Capture selection states AFTER Refresh
            // -------------------------------------------------
            List<bool> afterSelection = new List<bool>();
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                afterSelection.Add(item.Selected);
            }

            // Compare before and after states and output the result
            Console.WriteLine("Slicer selection state comparison:");
            for (int i = 0; i < beforeSelection.Count; i++)
            {
                bool before = beforeSelection[i];
                bool after = afterSelection[i];
                Console.WriteLine($"Item {i} (Value: {slicer.SlicerCache.SlicerCacheItems[i].Value}) - Before: {before}, After: {after}, Changed: {before != after}");
            }

            // Save the workbook (using the standard lifecycle rule)
            workbook.Save("SlicerRefreshComparisonDemo.xlsx");
        }
    }
}
