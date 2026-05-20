using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    public class SlicerRefreshDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet (data sheet)
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                Cells cells = dataSheet.Cells;

                // Populate initial source data (Product column and Sales column)
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Sales");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(100);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(150);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(200);

                // Add a worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create the pivot table using the source range A1:B4
                int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "D3", "ProductPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];

                // Set row field to Product and data field to Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
                pivotTable.CalculateData();

                // Add a slicer for the "Product" field on the pivot sheet
                int slicerIdx = pivotSheet.Slicers.Add(pivotTable, "F3", "Product");
                Slicer slicer = pivotSheet.Slicers[slicerIdx];

                // Record the initial number of items in the slicer cache (should be 3)
                int initialItemCount = slicer.SlicerCache.SlicerCacheItems.Count;
                Console.WriteLine("Initial slicer items count: " + initialItemCount);

                // ----- Add new items to the source data (which the slicer is based on) -----
                // Append new rows below the existing data
                cells["A5"].PutValue("Date");
                cells["B5"].PutValue(120);
                cells["A6"].PutValue("Elderberry");
                cells["B6"].PutValue(80);

                // Refresh the slicer – this also refreshes the associated pivot table
                slicer.Refresh();

                // Recalculate pivot data to reflect any changes
                pivotTable.CalculateData();

                // Verify that the slicer cache now contains the newly added items
                int updatedItemCount = slicer.SlicerCache.SlicerCacheItems.Count;
                Console.WriteLine("Updated slicer items count: " + updatedItemCount);

                // Simple confirmation that the pivot table reflects the new items
                int pivotRowItemCount = pivotTable.RowFields[0].PivotItems.Count;
                Console.WriteLine("Pivot row items count: " + pivotRowItemCount);

                // Save the workbook
                workbook.Save("SlicerRefreshDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SlicerRefreshDemo.Run();
        }
    }
}