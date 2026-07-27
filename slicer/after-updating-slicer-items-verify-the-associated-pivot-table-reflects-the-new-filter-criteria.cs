using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerPivotVerification
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data for the pivot table
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(150);
            cells["A4"].PutValue("Orange");
            cells["B4"].PutValue(200);
            cells["A5"].PutValue("Apple");
            cells["B5"].PutValue(80);
            cells["A6"].PutValue("Banana");
            cells["B6"].PutValue(70);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B6", "D3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Fruit" field
            int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // -------------------------------------------------
            // Update slicer items: select only "Apple" and deselect others
            // -------------------------------------------------
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                var item = slicer.SlicerCache.SlicerCacheItems[i];
                // The Value property holds the item caption (e.g., "Apple")
                if (item.Value.Equals("Apple", StringComparison.OrdinalIgnoreCase))
                {
                    item.Selected = true; // select Apple
                }
                else
                {
                    item.Selected = false; // deselect all others
                }
            }

            // Refresh the slicer – this also refreshes and recalculates the linked pivot table
            slicer.Refresh();

            // -------------------------------------------------
            // Verify that the pivot table now reflects the slicer filter
            // -------------------------------------------------
            // After the filter, the pivot table should contain only one row item ("Apple")
            int visibleRowItemCount = pivot.RowFields[0].PivotItems.Count;
            Console.WriteLine("Visible row items in pivot table after slicer refresh: " + visibleRowItemCount);

            // Additionally, print the visible item captions to confirm
            for (int i = 0; i < visibleRowItemCount; i++)
            {
                var pivotItem = pivot.RowFields[0].PivotItems[i];
                Console.WriteLine($"Pivot Item {i + 1}: {pivotItem.Value}");
            }

            // Save the workbook (using the provided lifecycle rule)
            workbook.Save("SlicerPivotVerification.xlsx");
        }
    }
}