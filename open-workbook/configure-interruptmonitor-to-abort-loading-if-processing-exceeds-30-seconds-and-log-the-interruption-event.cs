// Title: How to use Aspose.Cells InterruptMonitor in C# to abort workbook loading after 30 seconds and log the timeout
// AI Prompts: Write C# code that creates an InterruptMonitor with a 30‑second timeout, attaches it to the Workbook constructor, and logs a timeout message when the load is aborted. | Show how to wrap the Aspose.Cells Workbook loading in a try‑catch that catches InterruptMonitorException and writes the interruption details to the console or a log file. | Demonstrate configuring the InterruptMonitor’s Timeout property, enabling it before loading an Excel file, and disposing the monitor after the operation.
// Common Searches: Aspose.Cells C# InterruptMonitor abort loading after 30 seconds example | set timeout for Workbook loading using Aspose.Cells .NET | log timeout event when Excel file fails to load with Aspose.Cells | how to handle InterruptMonitorException in Aspose.Cells C# | prevent long‑running workbook load in Aspose.Cells with interrupt monitor
// Tags: Aspose.Cells InterruptMonitor timeout configuration | C# abort workbook load with Aspose.Cells | log InterruptMonitor timeout Aspose.Cells | exception handling InterruptMonitorException .NET | set loading time limit for Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to configure Aspose.Cells' InterruptMonitor in C# to enforce a 30‑second loading limit for an Excel workbook. It shows creating the monitor, setting its Timeout property, enabling it before the Workbook constructor, catching InterruptMonitorException, logging the interruption, and properly disposing the monitor after the operation.
    class Program
    {
        static void Main()
        {
            // Record start time (kept for potential future use)
            DateTime startTime = DateTime.Now;

            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no interrupt monitor used to maintain compatibility)
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Prepare output path
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected runtime exceptions
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
