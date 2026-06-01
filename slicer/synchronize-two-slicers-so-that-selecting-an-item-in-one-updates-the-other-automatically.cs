using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerSyncDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Year";
            cells["C1"].Value = "Amount";

            string[] fruits = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
            int[] years = { 2020, 2020, 2020, 2021, 2021, 2021 };
            int[] amounts = { 50, 70, 90, 60, 80, 100 };

            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 1, 0].Value = fruits[i];
                cells[i + 1, 1].Value = years[i];
                cells[i + 1, 2].Value = amounts[i];
            }

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:C7", "E3", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add first slicer for the "Fruit" field
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx1 = slicers.Add(pivot, "G1", "Fruit");
            Slicer slicer1 = slicers[slicerIdx1];
            slicer1.Caption = "Fruit Slicer 1";

            // Add second slicer for the same "Fruit" field
            int slicerIdx2 = slicers.Add(pivot, "G15", "Fruit");
            Slicer slicer2 = slicers[slicerIdx2];
            slicer2.Caption = "Fruit Slicer 2";

            // ---- Synchronization demonstration ----
            // Select "Banana" in the first slicer by manipulating its SlicerCache
            foreach (SlicerCacheItem item in slicer1.SlicerCache.SlicerCacheItems)
            {
                // Set Selected = true only for "Banana", false for others
                item.Selected = string.Equals(item.Value?.ToString(), "Banana", StringComparison.OrdinalIgnoreCase);
            }

            // Refresh both slicers so they reflect the same selection
            slicer1.Refresh();
            slicer2.Refresh();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("SlicerSyncDemo.xlsx");
        }
    }
}