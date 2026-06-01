using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsTests
{
    public class SlicerRefreshTests
    {
        public static void Main(string[] args)
        {
            try
            {
                RunTest();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        private static void RunTest()
        {
            try
            {
                // Create a new workbook and a data worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate source data
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("Apple");
                dataSheet.Cells["B2"].PutValue(100);
                dataSheet.Cells["A3"].PutValue("Banana");
                dataSheet.Cells["B3"].PutValue(200);

                // Add a worksheet for the pivot table and create the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                int pivotIdx = pivotSheet.PivotTables.Add("Data!A1:B3", "C3", "PivotTable1");
                PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as data field

                // Calculate initial pivot data
                pivot.CalculateData();

                // Capture the initial number of row items in the pivot table
                int initialRowCount = pivot.RowFields[0].PivotItems.Count;

                // Add a slicer linked to the pivot table (based on the Product field)
                int slicerIdx = pivotSheet.Slicers.Add(pivot, "A1", "Product");
                Slicer slicer = pivotSheet.Slicers[slicerIdx];

                // Modify the source data by adding a new distinct product
                dataSheet.Cells["A4"].PutValue("Orange");
                dataSheet.Cells["B4"].PutValue(150);

                // Refresh the slicer, which also refreshes the linked pivot table
                slicer.Refresh();

                // After refresh, the pivot table should contain one additional row item
                int refreshedRowCount = pivot.RowFields[0].PivotItems.Count;

                // Verify that the row count increased by exactly one
                if (refreshedRowCount != initialRowCount + 1)
                {
                    throw new InvalidOperationException(
                        $"Row count mismatch. Expected {initialRowCount + 1}, but got {refreshedRowCount}.");
                }

                // Save the workbook (optional for visual verification)
                string outputPath = "SlicerRefreshTest.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Propagate exception to the outer handler
                throw new ApplicationException("Test execution failed.", ex);
            }
        }
    }
}