// Title: C# – Log worksheet enumeration timestamps with AspNet.Cells for large workbook performance
// Description: Demonstrates how to open a massive Excel file using Aspose.Cells, record the start and end times of worksheet enumeration, log per‑sheet row/column counts, calculate elapsed milliseconds, and append the data to a performance log file while optionally saving a copy of the workbook.
// Keywords: Aspose.Cells worksheet enumeration timing | C# log Excel processing performance | measure workbook enumeration duration | SystemTimeInterruptMonitor usage | log large Excel file performance
// Common Searches: log worksheet enumeration start end time Aspose.Cells C# | measure Excel workbook processing time with Aspose | SystemTimeInterruptMonitor example for performance monitoring | append performance data to text file in C# | track large workbook enumeration duration
// Developer Intent: Record the start and end timestamps of worksheet enumeration to a log file for performance analysis of large Excel workbooks.
// Use Cases: Identify bottlenecks by measuring how long worksheet enumeration takes in massive workbooks. | Capture per‑sheet used rows and columns for reporting and capacity planning. | Maintain a persistent log across runs to monitor performance trends over time. | Combine enumeration timing with SystemTimeInterruptMonitor to avoid interruptions during long operations.
// AI Prompts: Create C# code that uses Aspose.Cells to enumerate all worksheets, log start/end timestamps, and write elapsed time to a text file. | Show how to integrate SystemTimeInterruptMonitor with LoadOptions for uninterrupted performance measurement in Aspose.Cells. | Provide an example that logs each sheet's used rows and columns while appending overall enumeration metrics to a persistent log.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open a massive Excel file using Aspose.Cells, record the start and end times of worksheet enumeration, log per‑sheet row/column counts, calculate elapsed milliseconds, and append the data to a performance log file while optionally saving a copy of the workbook.
class WorksheetEnumerationLogger
{
    static void Main()
    {
        // Paths for input, output and log files
        string inputPath = "LargeWorkbook.xlsx";
        string outputPath = "LargeWorkbook_Processed.xlsx";
        string logPath = "PerformanceLog.txt";

        // Verify that the input workbook exists before proceeding
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Open the log file in append mode
            using (StreamWriter logWriter = new StreamWriter(logPath, true))
            {
                logWriter.WriteLine("=== Worksheet Enumeration Log: {0} ===", DateTime.Now);

                // Prepare a SystemTimeInterruptMonitor (optional – used for precise timing)
                SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
                LoadOptions loadOptions = new LoadOptions
                {
                    InterruptMonitor = monitor
                };
                monitor.StartMonitor(int.MaxValue); // high limit to avoid interruption

                // Load the workbook with the monitor‑enabled options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Record enumeration start time
                DateTime enumStart = DateTime.Now;
                logWriter.WriteLine("Enumeration Start: {0:O}", enumStart);

                // Enumerate worksheets and perform a lightweight operation
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int usedRows = sheet.Cells.MaxDataRow + 1;
                    int usedCols = sheet.Cells.MaxDataColumn + 1;
                    logWriter.WriteLine("Sheet \"{0}\" - UsedRows: {1}, UsedCols: {2}",
                                        sheet.Name, usedRows, usedCols);
                }

                // Record enumeration end time and elapsed duration
                DateTime enumEnd = DateTime.Now;
                logWriter.WriteLine("Enumeration End:   {0:O}", enumEnd);
                TimeSpan elapsed = enumEnd - enumStart;
                logWriter.WriteLine("Total Enumeration Time: {0} ms", elapsed.TotalMilliseconds);
                logWriter.WriteLine(); // blank line for readability
            }

            // Save the workbook (optional – here we simply copy it)
            Workbook wbToSave = new Workbook(inputPath);
            wbToSave.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
