// Title: C# – Measure Aspose.Cells PivotTable refresh time and log RefreshDate
// Description: Creates a workbook, adds sample data, builds a PivotTable, then uses Stopwatch to time RefreshData and CalculateData calls. The elapsed milliseconds and the PivotTable.RefreshDate are written to the console before saving the file.
// Keywords: Aspose.Cells PivotTable refresh time | C# measure pivot cache duration | log PivotTable RefreshDate | Stopwatch Aspose.Cells performance | benchmark pivot table refresh .NET | Aspose.Cells RefreshData CalculateData timing
// Common Searches: how to time pivot table refresh Aspose.Cells C# | log pivot refresh duration and date with Aspose.Cells | measure performance of PivotTable.RefreshData in .NET | Aspose.Cells benchmark pivot cache refresh | C# code to capture PivotTable.RefreshDate
// Developer Intent: Determine how long a PivotTable refresh takes after loading its cache and record the refresh timestamp.
// Use Cases: Performance testing of large workbooks that contain PivotTables. | Auditing data freshness by storing the last refresh timestamp. | Automated regression tests that flag slow pivot refresh operations.
// AI Prompts: Generate C# code using Aspose.Cells that refreshes a PivotTable, measures the elapsed time with Stopwatch, and prints the duration and RefreshDate. | Show how to capture and log the RefreshDate property after calling RefreshData and CalculateData on a PivotTable. | Provide an example that benchmarks PivotTable refresh performance and outputs a formatted console message.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshTiming
{
    // Creates a workbook, adds sample data, builds a PivotTable, then uses Stopwatch to time RefreshData and CalculateData calls. The elapsed milliseconds and the PivotTable.RefreshDate are written to the console before saving the file.
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
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(950);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

            // Measure the time taken to refresh the pivot table
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Refresh the pivot cache and calculate the pivot data
            pivotTable.RefreshData();   // Gathers data from the source
            pivotTable.CalculateData(); // Calculates the pivot report

            sw.Stop();

            // Log the duration and the refresh date
            Console.WriteLine($"Pivot table refresh duration: {sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"Pivot table last refreshed on: {pivotTable.RefreshDate}");

            // Save the workbook (optional, just to demonstrate persistence)
            workbook.Save("PivotRefreshTimingDemo.xlsx");
        }
    }
}
