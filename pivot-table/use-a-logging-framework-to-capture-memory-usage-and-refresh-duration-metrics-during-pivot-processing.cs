using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotProcessing
{
    class PivotProcessingLogger
    {
        static void Main()
        {
            try
            {
                // Simple console logger
                void LogInfo(string message, params object[] args) => Console.WriteLine(message, args);
                void LogError(string message, params object[] args) => Console.Error.WriteLine(message, args);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(300);

                // Add a pivot table to the worksheet
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as data field

                // Measure memory usage and time before refresh
                long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
                Stopwatch sw = Stopwatch.StartNew();

                // Refresh pivot data and calculate results
                pivot.RefreshData();
                pivot.CalculateData();

                // Stop timer and capture memory after refresh
                sw.Stop();
                long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
                long memoryUsed = memoryAfter - memoryBefore;

                // Log the performance metrics
                LogInfo("Pivot refresh duration: {0} ms", sw.ElapsedMilliseconds);
                LogInfo("Memory used during refresh: {0} bytes", memoryUsed);

                // Save the workbook (output file)
                string outputPath = "PivotWithLogging.xlsx";
                workbook.Save(outputPath);
                LogInfo("Workbook saved to {0}", outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: {0}", ex.Message);
            }
        }
    }
}