// Title: Log memory consumption and refresh duration while processing an Aspose.Cells pivot table in C#
// AI Prompts: Write C# code that uses GC.GetTotalMemory and Stopwatch to record memory before and after PivotTable.RefreshData and logs the results with a console logger. | Show how to capture the elapsed time of PivotTable.CalculateData and include it in the same performance log. | Demonstrate adding a simple Action<string> logger to an Aspose.Cells workflow to output memory and timing metrics for pivot table operations.
// Common Searches: aspnet log memory usage during pivot table refresh Aspose.Cells | measure execution time of PivotTable.RefreshData in C# | how to benchmark Aspose.Cells pivot table performance | capture GC memory before and after RefreshData Aspose.Cells example | log pivot table calculation duration with Stopwatch in .NET
// Tags: Aspose.Cells pivot table performance logging | C# memory measurement GC Aspose.Cells | Stopwatch timing PivotTable.RefreshData | log pivot refresh duration .NET | track memory delta RefreshData Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills it with data, adds a pivot table, and uses GC.GetTotalMemory and Stopwatch to log memory consumption before and after RefreshData as well as the duration of RefreshData and CalculateData. All metrics are written via a simple console logger before the workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            // Simple logger that writes to console
            Action<string> LogInfo = message => Console.WriteLine(message);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(300);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

            // Capture memory usage before refreshing the pivot table
            long memoryBefore = GC.GetTotalMemory(true);
            LogInfo($"Memory before RefreshData: {memoryBefore / 1024} KB");

            // Measure the time taken to refresh the pivot table data
            Stopwatch sw = Stopwatch.StartNew();
            PivotRefreshState refreshState = pivot.RefreshData();
            sw.Stop();
            LogInfo($"RefreshData duration: {sw.ElapsedMilliseconds} ms, state: {refreshState}");

            // Capture memory usage after refreshing
            long memoryAfter = GC.GetTotalMemory(true);
            LogInfo($"Memory after RefreshData: {memoryAfter / 1024} KB");
            LogInfo($"Memory delta: {(memoryAfter - memoryBefore) / 1024} KB");

            // Optionally measure the time taken to calculate the pivot data
            sw.Restart();
            pivot.CalculateData();
            sw.Stop();
            LogInfo($"CalculateData duration: {sw.ElapsedMilliseconds} ms");

            // Save the workbook
            string outputPath = "PivotMetricsDemo.xlsx";
            workbook.Save(outputPath);
            LogInfo($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
