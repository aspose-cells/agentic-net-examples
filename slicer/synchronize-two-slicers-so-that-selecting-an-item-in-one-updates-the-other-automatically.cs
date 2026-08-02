using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Fruit";
                cells["B1"].Value = "Year";
                cells["C1"].Value = "Amount";

                string[] fruits = { "Apple", "Banana", "Cherry", "Date" };
                int[] years = { 2020, 2021 };
                int row = 1;
                foreach (var fruit in fruits)
                {
                    foreach (var year in years)
                    {
                        cells[row, 0].Value = fruit;
                        cells[row, 1].Value = year;
                        cells[row, 2].Value = (row + 1) * 10;
                        row++;
                    }
                }

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("A1:C9", "E3", "FruitPivot");
                PivotTable pivot = pivots[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add two slicers that filter the same pivot field ("Fruit")
                SlicerCollection slicers = sheet.Slicers;
                int slicerIdx1 = slicers.Add(pivot, "G1", "Fruit");
                int slicerIdx2 = slicers.Add(pivot, "G15", "Fruit");

                Slicer slicer1 = slicers[slicerIdx1];
                Slicer slicer2 = slicers[slicerIdx2];

                // Example: select the first item in slicer1 (e.g., "Apple")
                foreach (SlicerCacheItem item in slicer1.SlicerCache.SlicerCacheItems)
                {
                    // Reset all selections first
                    item.Selected = false;
                }
                if (slicer1.SlicerCache.SlicerCacheItems.Count > 0)
                {
                    slicer1.SlicerCache.SlicerCacheItems[0].Selected = true; // select first item
                }

                // Synchronize slicer2 with slicer1
                for (int i = 0; i < slicer1.SlicerCache.SlicerCacheItems.Count; i++)
                {
                    SlicerCacheItem srcItem = slicer1.SlicerCache.SlicerCacheItems[i];
                    // Find matching item in slicer2 by value
                    for (int j = 0; j < slicer2.SlicerCache.SlicerCacheItems.Count; j++)
                    {
                        SlicerCacheItem dstItem = slicer2.SlicerCache.SlicerCacheItems[j];
                        if (srcItem.Value.Equals(dstItem.Value))
                        {
                            dstItem.Selected = srcItem.Selected;
                            break;
                        }
                    }
                }

                // Refresh both slicers so the UI reflects the synchronized state
                slicer1.Refresh();
                slicer2.Refresh();

                // Save the workbook
                workbook.Save("SynchronizedSlicers.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}