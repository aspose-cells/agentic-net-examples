using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCsvSelection
{
    class Program
    {
        static void Main()
        {
            // ---------- Create workbook and sample data ----------
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Sample data for pivot table (Fruit, Sales)
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(150);
            cells["A4"].PutValue("Orange");
            cells["B4"].PutValue(90);
            cells["A5"].PutValue("Grape");
            cells["B5"].PutValue(60);

            // ---------- Create pivot table ----------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIdx = pivotSheet.PivotTables.Add("A1:B5", "C3", "FruitPivot");
            PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add slicer linked to the pivot ----------
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIdx = slicerSheet.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer = slicerSheet.Slicers[slicerIdx];

            // ---------- Load external CSV containing values to select ----------
            // Expected CSV format: one value per line, e.g.
            // Apple
            // Orange
            string csvPath = "selection.csv";
            HashSet<string> valuesToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(csvPath))
            {
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        valuesToSelect.Add(trimmed);
                }
            }
            else
            {
                Console.WriteLine($"CSV file '{csvPath}' not found. No items will be selected.");
            }

            // ---------- Update slicer cache items based on CSV ----------
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Deselect all items first
                item.Selected = false;
                // Select if the item's value exists in the CSV list
                if (valuesToSelect.Contains(item.Value))
                {
                    item.Selected = true;
                }
            }

            // ---------- Refresh slicer to apply selection ----------
            slicer.Refresh();

            // ---------- Save the workbook ----------
            workbook.Save("SlicerWithCsvSelection.xlsx");
        }
    }
}