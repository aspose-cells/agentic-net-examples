// Title: C# example to capture and compare Aspose.Cells slicer item selection states before and after invoking Refresh
// Common Searches: Aspose.Cells C# compare slicer selected items before and after Refresh | how to verify slicer selection changes with Refresh method in Aspose.Cells | C# code to read SlicerCacheItem Selected property in Aspose.Cells | Aspose.Cells slicer Refresh not applying selection changes example | capture initial slicer state and compare after Refresh using Aspose.Cells .NET | Aspose.Cells pivot table slicer selection state validation C#
// Tags: Aspose.Cells slicer selection comparison | C# slicer Refresh verification | Aspose.Cells pivot table slicer cache items | Aspose.Cells capture slicer selected property | Aspose.Cells workbook slicer state validation

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRefreshComparison
{
    // The sample creates a workbook, adds sample data, builds a pivot table, and inserts a slicer linked to the "Fruit" field. It records the initial Selected state of each slicer item, changes the selection to keep only the first item, calls slicer.Refresh(), captures the post‑refresh states, compares the two lists, outputs any differences, and saves the workbook as SlicerRefreshComparison.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Populate sample data for a pivot table
                // ------------------------------------------------------------
                // Header row
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                // Data rows
                string[] fruits = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
                int[] years = { 2020, 2020, 2020, 2021, 2021, 2021, 2022, 2022, 2022 };
                int[] amounts = { 50, 60, 70, 80, 90, 100, 110, 120, 130 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 1, 0].PutValue(fruits[i]);   // Column A
                    cells[i + 1, 1].PutValue(years[i]);   // Column B
                    cells[i + 1, 2].PutValue(amounts[i]); // Column C
                }

                // ------------------------------------------------------------
                // 2. Create a pivot table based on the data range
                // ------------------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("=Sheet1!A1:C9", "E12", "FruitPivot");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // ------------------------------------------------------------
                // 3. Add a slicer linked to the "Fruit" field of the pivot table
                // ------------------------------------------------------------
                // Note: The Add method expects destination cell first, then the field name.
                int slicerIdx = sheet.Slicers.Add(pivot, "F2", "Fruit");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // ------------------------------------------------------------
                // 4. Capture the initial selection states of all slicer items
                // ------------------------------------------------------------
                List<bool> beforeSelection = CaptureSelectionStates(slicer);
                Console.WriteLine("Initial selection states:");
                PrintSelectionStates(slicer, beforeSelection);

                // ------------------------------------------------------------
                // 5. Modify selection: select only the first item, deselect the rest
                // ------------------------------------------------------------
                SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
                for (int i = 0; i < items.Count; i++)
                {
                    items[i].Selected = (i == 0); // true for first item, false otherwise
                }

                // ------------------------------------------------------------
                // 6. Refresh the slicer to apply the selection changes
                // ------------------------------------------------------------
                slicer.Refresh();

                // ------------------------------------------------------------
                // 7. Capture the selection states after refresh
                // ------------------------------------------------------------
                List<bool> afterSelection = CaptureSelectionStates(slicer);
                Console.WriteLine("\nSelection states after Refresh:");
                PrintSelectionStates(slicer, afterSelection);

                // ------------------------------------------------------------
                // 8. Compare before and after states and report differences
                // ------------------------------------------------------------
                Console.WriteLine("\nComparison result:");
                bool anyDifference = false;
                for (int i = 0; i < beforeSelection.Count; i++)
                {
                    if (beforeSelection[i] != afterSelection[i])
                    {
                        anyDifference = true;
                        Console.WriteLine($"Item \"{items[i].Value}\": before={beforeSelection[i]}, after={afterSelection[i]}");
                    }
                }

                if (!anyDifference)
                {
                    Console.WriteLine("No differences detected – Refresh did not apply the changes.");
                }

                // ------------------------------------------------------------
                // 9. Save the workbook (optional, for visual verification)
                // ------------------------------------------------------------
                workbook.Save("SlicerRefreshComparison.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static List<bool> CaptureSelectionStates(Slicer slicer)
        {
            List<bool> states = new List<bool>();
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
            foreach (SlicerCacheItem item in items)
            {
                states.Add(item.Selected);
            }
            return states;
        }

        private static void PrintSelectionStates(Slicer slicer, List<bool> states)
        {
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < items.Count; i++)
            {
                Console.WriteLine($"  Item \"{items[i].Value}\": Selected = {states[i]}");
            }
        }
    }
}
