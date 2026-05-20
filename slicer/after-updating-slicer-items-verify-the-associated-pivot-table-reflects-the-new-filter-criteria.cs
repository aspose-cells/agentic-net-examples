using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerPivotDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source data (Fruit | Sales)
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Sales");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(80);
                cells["A4"].PutValue("Apple");
                cells["B4"].PutValue(150);
                cells["A5"].PutValue("Orange");
                cells["B5"].PutValue(200);

                // Add a pivot table based on the source data
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");   // Row field
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales"); // Data field (Sum)

                // Refresh data and calculate initial results
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Fruit" field
                int slicerIdx = sheet.Slicers.Add(pivot, "G3", "Fruit");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // ---- Update slicer items: select only "Apple" ----
                for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
                {
                    SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
                    // Select the item if its value is "Apple", otherwise deselect
                    item.Selected = string.Equals(item.Value, "Apple", StringComparison.OrdinalIgnoreCase);
                }

                // Refresh the slicer – this also refreshes and recalculates the linked pivot table
                slicer.Refresh();

                // ---- Verify pivot table reflects the slicer filter ----
                // After the slicer filter, the sum for Apple will be in cell E4 (first data row, second column of the pivot)
                double appleSum = 0;
                try
                {
                    Cell sumCell = sheet.Cells["E4"];
                    if (sumCell.Value != null && double.TryParse(sumCell.Value.ToString(), out double d))
                    {
                        appleSum = d;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read pivot result: {ex.Message}");
                }

                Console.WriteLine($"Sum of Sales for Apple after slicer filter: {appleSum}");

                // Save the workbook (ensure the path is valid)
                string outputPath = "SlicerPivotRefreshDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}