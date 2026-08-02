using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSelectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(300);

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table (field "Fruit")
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // OPTIONAL: Deselect all items to simulate a slicer with no selection
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = false;
            }

            // BEFORE saving: ensure at least one item is selected.
            // If none are selected, select the first item.
            bool anySelected = false;
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                if (item.Selected)
                {
                    anySelected = true;
                    break;
                }
            }

            if (!anySelected && slicer.SlicerCache.SlicerCacheItems.Count > 0)
            {
                // Select the first slicer item
                slicer.SlicerCache.SlicerCacheItems[0].Selected = true;
            }

            // Save the workbook
            workbook.Save("SlicerSelectionDemo.xlsx");
        }
    }
}