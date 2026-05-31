using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerAutoSelectDemo
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
            cells["A2"].PutValue("Apple");
            cells["A3"].PutValue("Orange");
            cells["A4"].PutValue("Banana");
            cells["B1"].PutValue("Sales");
            cells["B2"].PutValue(100);
            cells["B3"].PutValue(200);
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

            // Refresh the slicer to ensure cache is up‑to‑date
            slicer.Refresh();

            // BEFORE saving: ensure each slicer has at least one selected item
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Slicer s in ws.Slicers)
                {
                    SlicerCacheItemCollection items = s.SlicerCache.SlicerCacheItems;
                    bool anySelected = false;

                    // Check if any item is already selected
                    for (int i = 0; i < items.Count; i++)
                    {
                        if (items[i].Selected)
                        {
                            anySelected = true;
                            break;
                        }
                    }

                    // If none selected, select the first item
                    if (!anySelected && items.Count > 0)
                    {
                        items[0].Selected = true;
                        // Optionally set the first visible item index
                        s.FirstItemIndex = 0;
                    }
                }
            }

            // Save the workbook
            workbook.Save("SlicerAutoSelectDemo.xlsx");
        }
    }
}