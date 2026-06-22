using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceMonitoring
{
    class Program
    {
        static void Main()
        {
            // Path to the large workbook to be processed
            const string inputPath = "Large.xlsx";
            // Path where the processed workbook will be saved
            const string outputPath = "Large_processed.xlsx";
            // Path to the performance log file
            const string logPath = "performance.log";

            // Ensure the log file exists and open it for appending
            using (StreamWriter logWriter = new StreamWriter(logPath, true))
            {
                try
                {
                    // Create a SystemTimeInterruptMonitor (no interruption limit needed for logging)
                    SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
                    LoadOptions loadOptions = new LoadOptions
                    {
                        InterruptMonitor = monitor
                    };

                    // Start the monitor (use a very large limit so it does not interrupt)
                    monitor.StartMonitor(int.MaxValue);

                    // Log the start timestamp
                    DateTime startTime = DateTime.Now;
                    logWriter.WriteLine($"Enumeration started at: {startTime:O}");

                    // Load the large workbook with the monitor attached
                    Workbook workbook = new Workbook(inputPath, loadOptions);

                    // Enumerate each worksheet and optionally iterate through its rows
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Example enumeration: iterate through used rows in column A
                        Cells cells = sheet.Cells;
                        int maxDataRow = cells.MaxDataRow;
                        for (int row = 0; row <= maxDataRow; row++)
                        {
                            // Access a cell to simulate work (no modification needed)
                            _ = cells[row, 0].Value;
                        }
                    }

                    // Log the end timestamp
                    DateTime endTime = DateTime.Now;
                    logWriter.WriteLine($"Enumeration ended at: {endTime:O}");
                    logWriter.WriteLine($"Total duration (seconds): {(endTime - startTime).TotalSeconds:F2}");
                    logWriter.WriteLine(new string('-', 40));
                    logWriter.Flush();

                    // Save the workbook after processing
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    // Log any exception that occurs during processing
                    logWriter.WriteLine($"Error: {ex.Message}");
                    logWriter.WriteLine($"StackTrace: {ex.StackTrace}");
                }
            }
        }
    }
}