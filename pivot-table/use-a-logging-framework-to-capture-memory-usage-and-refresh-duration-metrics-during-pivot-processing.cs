// Title: Log Pivot Table Refresh Time and Memory Usage with Aspose.Cells for .NET
// Description: Demonstrates how to measure and log the execution time and memory consumption of a PivotTable refresh in Aspose.Cells using Stopwatch and GC.GetTotalMemory, with examples for console output and integration with structured logging frameworks.
// Keywords: Aspose.Cells | PivotTable performance | C# memory profiling | Stopwatch timing | GC.GetTotalMemory | logging framework | Serilog | NLog | Microsoft.Extensions.Logging | performance metrics
// Common Searches: measure pivot table refresh time Aspose.Cells C# | log memory usage during pivot refresh .NET | Aspose.Cells performance logging example | how to profile pivot table calculation in C# | track execution time of Aspose.Cells PivotTable
// Developer Intent: Capture execution duration and memory delta of a PivotTable refresh and persist the data with a logging solution.
// Use Cases: Replace Console.WriteLine with a structured logger (e.g., Serilog, NLog) to store refresh metrics in files or monitoring systems. | Create a reusable helper method that accepts a PivotTable, records time and memory, logs the results, and returns a metric object for benchmarking. | Integrate the logging routine into automated tests to detect performance regressions across different data sets or pivot configurations.
// AI Prompts: Generate C# code that uses Serilog to log pivot refresh duration and memory usage measured with Stopwatch and GC.GetTotalMemory. | Show a method signature that takes a PivotTable, records memory before/after RefreshData and CalculateData, logs the metrics via Microsoft.Extensions.Logging, and returns a custom PerformanceResult object. | Provide an example that writes timestamp, duration, and memory delta to a CSV file after each pivot refresh using Aspose.Cells.

using System;
using System.Diagnostics;

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to measure and log the execution time and memory consumption of a PivotTable refresh in Aspose.Cells using Stopwatch and GC.GetTotalMemory, with examples for console output and integration with structured logging frameworks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);

            // Capture memory usage before refresh
            long memoryBefore = GC.GetTotalMemory(true);
            Stopwatch sw = Stopwatch.StartNew();

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Stop timer and capture memory after refresh
            sw.Stop();
            long memoryAfter = GC.GetTotalMemory(true);
            long memoryUsed = memoryAfter - memoryBefore;

            // Log duration and memory consumption
            Console.WriteLine("Pivot refresh duration: {0} ms", sw.ElapsedMilliseconds);
            Console.WriteLine("Memory used during refresh: {0} bytes", memoryUsed);

            // Save the workbook
            string outputPath = "PivotWithMetrics.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
