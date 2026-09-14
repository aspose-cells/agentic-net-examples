// Title: How to measure and log the refresh duration of each DBConnection in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, enumerates all DBConnection objects, toggles BackgroundRefresh to trigger a refresh, and captures the elapsed milliseconds using Stopwatch. | Show how to output the refresh time for each DBConnection to the console in a readable format while handling missing input files gracefully. | Explain how to save the workbook after measuring DBConnection refresh performance and ensure any changes are persisted.
// Common Searches: c# Aspose.Cells how to benchmark the refresh time of external DB connections in an Excel workbook | measure performance of DataConnections using Stopwatch in Aspose.Cells .NET | log refresh latency of each DBConnection object while processing a workbook with Aspose.Cells
// Tags: measure DBConnection refresh Aspose.Cells | log external connection latency .NET | benchmark Excel data connections C# | stopwatch timing DBConnection Aspose.Cells | performance profiling Aspose.Cells external connections

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // // Loads an Excel workbook, iterates through its DBConnection objects, measures the time taken to refresh each connection by toggling BackgroundRefresh, logs the elapsed milliseconds to the console, and saves the workbook.
    public class DBConnectionRefreshPerformanceDemo
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing external DB connections
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through each connection and measure refresh time for DBConnection objects
                foreach (ExternalConnection conn in connections)
                {
                    if (conn is DBConnection dbConn)
                    {
                        Console.WriteLine($"Processing DBConnection: {dbConn.Name}");

                        // Start timing
                        Stopwatch sw = Stopwatch.StartNew();

                        // Simulate a refresh by toggling a property that forces re‑evaluation
                        bool originalBackgroundRefresh = dbConn.BackgroundRefresh;
                        dbConn.BackgroundRefresh = !originalBackgroundRefresh;
                        dbConn.BackgroundRefresh = originalBackgroundRefresh;

                        // Stop timing
                        sw.Stop();

                        // Log the elapsed time
                        Console.WriteLine($"Refresh time for '{dbConn.Name}': {sw.ElapsedMilliseconds} ms");
                    }
                }

                // Save the workbook (if any changes were made)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
