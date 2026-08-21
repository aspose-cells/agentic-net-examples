// Title: Capture Memory Usage and Execution Time for Aspose.Cells Pivot Table Refresh in C# (.NET)
// Description: This Aspose.Cells .NET example creates a workbook, adds a simple pivot table, and uses GC.GetTotalMemory together with Stopwatch to record memory consumption and elapsed time during RefreshData and CalculateData. The metrics are logged via a placeholder console logger (easily replaceable with NLog, Serilog, or any logging framework) before the workbook is saved.
// Keywords: Aspose.Cells | pivot table | C# | .NET | performance logging | memory profiling | execution time | RefreshData | CalculateData | NLog | Serilog | GitHub example | sample code
// Common Searches: Aspose.Cells log pivot refresh time | measure memory usage during pivot table calculation .NET | how to profile Aspose.Cells pivot performance | C# example for timing pivot RefreshData | replace console logger with NLog in Aspose.Cells sample
// Developer Intent: Record and log memory delta and duration while refreshing and calculating a pivot table using Aspose.Cells for .NET.
// Use Cases: Benchmark pivot refresh speed for large datasets. | Integrate pivot processing metrics into existing monitoring or logging pipelines. | Compare memory impact of different pivot configurations. | Generate performance reports for data‑intensive Excel automation.
// AI Prompts: Show how to swap the console logger with NLog or Serilog for persistent metric storage. | Create a reusable helper method that returns an object containing duration and memory used for any pivot operation. | Demonstrate logging the captured metrics to a JSON file and uploading them to a monitoring dashboard.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This Aspose.Cells .NET example creates a workbook, adds a simple pivot table, and uses GC.GetTotalMemory together with Stopwatch to record memory consumption and elapsed time during RefreshData and CalculateData. The metrics are logged via a placeholder console logger (easily replaceable with NLog, Serilog, or any logging framework) before the workbook is saved.
class PivotProcessingMetrics
{
    static void Main()
    {
        try
        {
            // Simple console logger (replace NLog)
            Action<string> LogInfo = message => Console.WriteLine($"INFO: {message}");
            Action<string> LogError = message => Console.Error.WriteLine($"ERROR: {message}");

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
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);

            // Capture memory usage before refresh
            long memBefore = GC.GetTotalMemory(true);
            Stopwatch sw = Stopwatch.StartNew();

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Stop timing and capture memory after refresh
            sw.Stop();
            long memAfter = GC.GetTotalMemory(true);
            long memUsed = memAfter - memBefore;

            // Log metrics
            LogInfo("Pivot refresh and calculation completed.");
            LogInfo($"Duration: {sw.ElapsedMilliseconds} ms");
            LogInfo($"Memory used: {memUsed / 1024} KB");

            // Save the workbook
            string outputPath = "PivotWithMetrics.xlsx";
            workbook.Save(outputPath);
            LogInfo($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
