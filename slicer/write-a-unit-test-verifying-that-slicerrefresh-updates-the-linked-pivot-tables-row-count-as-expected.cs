// Title: C# Unit Test – Verify Slicer.Refresh Increments Pivot Table Row Count in Aspose.Cells
// Description: Creates a workbook with product data, builds a pivot table, attaches a slicer to the "Product" field, adds a new row to the source sheet, calls Slicer.Refresh, and asserts that the pivot table's row‑field item count grows by one.
// Keywords: Aspose.Cells | C# | .NET | Slicer.Refresh | pivot table unit test | row items count | Excel automation testing | MSTest example | NUnit pivot slicer | xUnit Aspose.Cells | GitHub sample code
// Common Searches: how to test slicer refresh in Aspose.Cells | unit test for pivot table row count after adding data | Aspose.Cells C# example verifying slicer‑pivot synchronization | MSTest NUnit xUnit slicer refresh test | Aspose.Cells slicer linked pivot table unit test
// Developer Intent: Confirm that invoking Slicer.Refresh correctly updates the associated pivot table's row items.
// Use Cases: Automated regression testing of dashboard components that rely on slicer‑driven pivot tables. | Continuous‑integration validation of dynamic Excel reports generated with Aspose.Cells. | Ensuring newly added source rows are reflected in slicer‑filtered pivot tables before release.
// AI Prompts: Generate an MSTest method that builds a workbook, adds a pivot table and slicer, inserts a new data row, calls Slicer.Refresh, and asserts the row count increased. | Write a NUnit test for Aspose.Cells that verifies Slicer.Refresh updates the linked pivot table's row field items. | Provide an xUnit example checking that a slicer refresh reflects added source data in the pivot table row count.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsTests
{
    // Creates a workbook with product data, builds a pivot table, attaches a slicer to the "Product" field, adds a new row to the source sheet, calls Slicer.Refresh, and asserts that the pivot table's row‑field item count grows by one.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Sales");

                // Initial rows
                dataSheet.Cells["A2"].PutValue("Apple");
                dataSheet.Cells["B2"].PutValue(100);
                dataSheet.Cells["A3"].PutValue("Banana");
                dataSheet.Cells["B3"].PutValue(200);

                // Add a worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "C3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot table: Product as row field, Sales as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

                // Refresh to calculate initial data
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table on the "Product" field
                int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "E3", "Product");
                Slicer slicer = pivotSheet.Slicers[slicerIndex];

                // Capture the initial row count of the pivot table
                int initialRowCount = pivotTable.RowFields[0].PivotItems.Count;

                // Add a new product row to the source data
                dataSheet.Cells["A4"].PutValue("Orange");
                dataSheet.Cells["B4"].PutValue(150);

                // Refresh via slicer; this should refresh the linked pivot table
                slicer.Refresh();

                // Capture the row count after refresh
                int refreshedRowCount = pivotTable.RowFields[0].PivotItems.Count;

                // Verify that the row count increased by one
                if (refreshedRowCount == initialRowCount + 1)
                {
                    Console.WriteLine("Test passed: Row count increased by 1 after slicer refresh.");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected row count {initialRowCount + 1}, but got {refreshedRowCount}.");
                }

                // Optional: Save the workbook for manual inspection
                // workbook.Save("SlicerRefreshTest.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
