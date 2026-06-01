using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerReset
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
                cells["A1"].Value = "Category";
                cells["A2"].Value = "A";
                cells["A3"].Value = "B";
                cells["A4"].Value = "C";
                cells["B1"].Value = "Amount";
                cells["B2"].Value = 10;
                cells["B3"].Value = 20;
                cells["B4"].Value = 30;

                // Add a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table
                // Note: The correct parameter order is (pivot, destinationCell, baseFieldName)
                SlicerCollection slicers = sheet.Slicers;
                int slicerIdx = slicers.Add(pivot, "E1", "Category");
                Slicer slicer = slicers[slicerIdx];

                // ---- USER FILTER SIMULATION ----
                // Select only the first item to simulate a filtered state
                SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
                for (int i = 0; i < items.Count; i++)
                {
                    items[i].Selected = i == 0; // only first item selected
                }
                slicer.Refresh();

                // ---- CLEAR ALL SELECTIONS (RESET FILTER) ----
                // To reset the slicer, select all items (equivalent to "Clear Filter")
                foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                {
                    item.Selected = true; // select every item
                }
                slicer.Refresh();

                // Save the workbook
                string outputPath = "SlicerResetDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}