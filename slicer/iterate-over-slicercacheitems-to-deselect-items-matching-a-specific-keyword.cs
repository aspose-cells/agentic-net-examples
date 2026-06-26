using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
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

                // Populate sample data for a pivot table
                cells["A1"].Value = "Fruit";
                cells["A2"].Value = "Apple";
                cells["A3"].Value = "Orange";
                cells["A4"].Value = "Banana";
                cells["A5"].Value = "Apple";

                cells["B1"].Value = "Quantity";
                cells["B2"].Value = 10;
                cells["B3"].Value = 20;
                cells["B4"].Value = 15;
                cells["B5"].Value = 12;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivot.RefreshData();

                // Add a slicer linked to the "Fruit" field of the pivot table.
                // Destination cell for the slicer (e.g., "E1") must be a valid cell address.
                SlicerCollection slicers = sheet.Slicers;
                int slicerIdx = slicers.Add(pivot, "E1", "Fruit");
                Slicer slicer = slicers[slicerIdx];
                slicer.Name = "FruitSlicer";
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

                // Keyword to deselect (e.g., deselect all "Apple" items)
                string keyword = "Apple";

                // Iterate over slicer cache items and set selection state
                SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
                foreach (SlicerCacheItem item in cacheItems)
                {
                    // Deselect items matching the keyword; select others
                    item.Selected = item.Value != keyword;
                }

                // Apply the changes to the slicer
                slicer.Refresh();

                // Save the workbook
                string outputPath = "SlicerDeselectKeywordDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}