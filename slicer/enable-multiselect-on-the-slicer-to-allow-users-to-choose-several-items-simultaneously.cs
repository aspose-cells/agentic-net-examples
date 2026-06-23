using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerMultiSelectDemo
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
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 150;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 90;
            cells["A5"].Value = "Apple";
            cells["B5"].Value = 80;
            cells["A6"].Value = "Orange";
            cells["B6"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Row field (Fruit) and data field (Sales)
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Enable multiple item selection for the page field (if any)
            // In this case the row field acts as a page field when a slicer is attached.
            // The property must be set on the corresponding PivotField.
            // Since we have only one row field, it is at index 0 of RowFields.
            // However, the IsMultipleItemSelectionAllowed property is relevant for PageFields.
            // To ensure slicer allows multi‑select, set it on the underlying PivotField.
            // First, add the row field also to PageFields (optional, but ensures the property works).
            pivot.PageFields.Add(pivot.RowFields[0]);
            PivotField pageField = pivot.PageFields[0];
            pageField.IsMultipleItemSelectionAllowed = true;

            // Add a slicer linked to the "Fruit" field
            int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
            slicer.Caption = "Fruit Slicer";

            // Select multiple items in the slicer (e.g., Apple and Orange)
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
                if (item.Value == "Apple" || item.Value == "Orange")
                {
                    item.Selected = true; // select the item
                }
                else
                {
                    item.Selected = false; // deselect other items
                }
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerMultiSelectDemo.xlsx");
        }
    }
}