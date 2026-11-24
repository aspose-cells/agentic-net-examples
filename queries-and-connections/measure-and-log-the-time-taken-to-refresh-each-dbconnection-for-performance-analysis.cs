// Title: Measure and log the time taken to refresh all external data connections in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to profile workbook.RefreshAll execution time with Aspose.Cells in C# | c# Aspose.Cells log refresh duration of external data connections to a file | measure performance of RefreshAll for DB connections using Stopwatch in .NET | track time taken to refresh Excel data connections with Aspose.Cells library | record elapsed milliseconds of workbook.RefreshAll in a log file
// Tags: Aspose.Cells RefreshAll execution timing | C# Stopwatch profiling of external data connections | write refresh elapsed time to log file | performance analysis of Excel data connections using Aspose.Cells | benchmark workbook.RefreshAll in .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses a System.Diagnostics.Stopwatch to measure how long workbook.RefreshAll takes to refresh all external DB and OLEDB connections, writes the elapsed milliseconds to both the console and a text log, then saves the workbook. Error handling ensures missing files or refresh failures are reported.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare a log file for performance data
            string logFilePath = "RefreshLog.txt";
            using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
            {
                // Refresh all external data connections in the workbook
                Stopwatch sw = Stopwatch.StartNew();
                try
                {
                    workbook.RefreshAll(); // Refreshes all DB and OLEDB connections
                }
                catch (Exception refreshEx)
                {
                    Console.WriteLine($"Error during refresh: {refreshEx.Message}");
                    logWriter.WriteLine($"Error during refresh: {refreshEx.Message}");
                }
                sw.Stop();

                string logEntry = $"All connections refreshed in {sw.ElapsedMilliseconds} ms.";
                Console.WriteLine(logEntry);
                logWriter.WriteLine(logEntry);
            }

            // Save the workbook after refreshing connections
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
