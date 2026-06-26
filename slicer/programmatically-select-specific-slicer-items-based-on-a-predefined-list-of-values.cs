using System;
using System.Collections.Generic;
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

            // Populate sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Orange";
            cells["A4"].Value = "Banana";
            cells["A5"].Value = "Apple";
            cells["B1"].Value = "Quantity";
            cells["B2"].Value = 10;
            cells["B3"].Value = 20;
            cells["B4"].Value = 30;
            cells["B5"].Value = 15;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);     // Quantity column
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Fruit" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Predefined list of slicer values that should be selected
            List<string> valuesToSelect = new List<string> { "Apple", "Banana" };

            // Iterate through slicer cache items and set selection based on the list
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Select the item if its value is in the predefined list; otherwise deselect
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // Refresh the slicer to apply the changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerSelectionDemo.xlsx");
        }
    }
}