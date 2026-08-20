// Title: Abort Aspose.Cells Workbook Load with SystemTimeInterruptMonitor Timeout (C#)
// Description: Demonstrates how to use SystemTimeInterruptMonitor with LoadOptions to stop a workbook load if it exceeds a defined time (e.g., 2000 ms). The monitor throws an exception, which can be caught to handle the abort gracefully.
// Keywords: Aspose.Cells | SystemTimeInterruptMonitor | Workbook load timeout | C# Excel loading | LoadOptions.InterruptMonitor | cancel large Excel file load | Aspose.Cells performance | .NET Excel processing
// Common Searches: Aspose.Cells stop workbook loading after timeout | SystemTimeInterruptMonitor example C# | How to abort large Excel file load with Aspose.Cells | Set time limit for Workbook loading Aspose.Cells .NET | InterruptMonitor usage in Aspose.Cells
// Developer Intent: Implement a time‑bound workbook loading operation that aborts automatically when the specified limit is exceeded.
// Use Cases: Prevent UI freeze by limiting Excel load time in desktop applications. | Enforce request‑level timeouts for Excel processing in web APIs. | Provide a quick preview of a massive workbook when full load exceeds the allowed duration.
// AI Prompts: Write C# code that loads an Excel file with SystemTimeInterruptMonitor set to a 5‑second timeout and logs a custom error message on abort. | Show how to retry workbook loading with a larger timeout after catching an interrupt exception. | Explain configuring LoadOptions.InterruptMonitor for asynchronous workbook loading in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to use SystemTimeInterruptMonitor with LoadOptions to stop a workbook load if it exceeds a defined time (e.g., 2000 ms). The monitor throws an exception, which can be caught to handle the abort gracefully.
    public class InterruptMonitorLoadingDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the workbook file
            const string filePath = "LargeFile.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Create a SystemTimeInterruptMonitor; false => throws exception when interrupted
                SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

                // Set up load options with the monitor
                LoadOptions loadOptions = new LoadOptions
                {
                    InterruptMonitor = monitor
                };

                // Define a time limit (e.g., 2000 ms) for the loading operation
                monitor.StartMonitor(2000);

                // Load the workbook using the load options that contain the monitor
                Workbook wb = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Loading was aborted because it exceeded the time limit or another error occurred
                Console.WriteLine("Loading aborted: " + ex.Message);
            }
        }
    }
}
