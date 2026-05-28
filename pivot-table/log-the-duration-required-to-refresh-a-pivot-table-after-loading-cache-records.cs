using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDuration
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["A5"].PutValue("Apple");
            sheet.Cells["B5"].PutValue(130);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

            // Initial refresh to ensure the pivot cache is built
            pivot.RefreshData();
            pivot.CalculateData();

            // Modify source data to simulate a change that requires refresh
            sheet.Cells["B2"].PutValue(200); // Change Apple sales

            // Measure the time taken to refresh the pivot table after the data change
            Stopwatch sw = Stopwatch.StartNew();
            pivot.RefreshData();      // Refresh cache from the data source
            pivot.CalculateData();    // Recalculate the pivot table values
            sw.Stop();

            // Log the duration
            Console.WriteLine($"Pivot table refresh duration: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (using the standard Aspose.Cells save method)
            workbook.Save("PivotRefreshDuration.xlsx");
        }
    }
}