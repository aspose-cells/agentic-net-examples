using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            try
            {
                RunSlicerRefreshTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void RunSlicerRefreshTest()
        {
            // Create a new workbook and populate source data
            var workbook = new Workbook();
            var dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(200);

            // Add a pivot table based on the source data
            var pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "C3", "PivotTable1");
            var pivotTable = pivotSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field
            pivotTable.CalculateData();

            // Add a slicer linked to the pivot table for the "Product" field
            var slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
            var slicer = slicerSheet.Slicers[slicerIndex];

            // Capture initial row count of the pivot table
            int initialRowCount = pivotTable.RowFields[0].PivotItems.Count;

            // Add a new product row to the source data
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(150);

            // Refresh via slicer – this should refresh the linked pivot table
            slicer.Refresh();

            // Capture the row count after refresh
            int refreshedRowCount = pivotTable.RowFields[0].PivotItems.Count;

            // Verify that the row count increased by one
            if (refreshedRowCount != initialRowCount + 1)
            {
                throw new InvalidOperationException(
                    $"Expected row count to increase by 1 after slicer refresh. " +
                    $"Initial: {initialRowCount}, After: {refreshedRowCount}");
            }

            // Optional: Save the workbook for manual inspection (ensure directory exists)
            string outputPath = "SlicerRefreshResult.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                // If saving fails, log but do not treat as test failure
                Console.WriteLine($"Warning: Could not save workbook. {saveEx.Message}");
            }
        }
    }
}