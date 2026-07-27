// Title: Abort Aspose.Cells workbook load after 5 seconds with SystemTimeInterruptMonitor (C#)
// Description: Demonstrates how to create a SystemTimeInterruptMonitor, attach it to LoadOptions, start a 5‑second timer, and load a workbook. If loading exceeds the limit, the monitor throws an exception that is caught and handled gracefully.
// Keywords: Aspose.Cells | SystemTimeInterruptMonitor | workbook load timeout | C# LoadOptions | interrupt monitor exception | Excel file loading abort | 5 second timeout | LoadOptions InterruptMonitor
// Common Searches: Aspose.Cells set timeout for workbook loading | How to cancel workbook load after specific time in .NET | SystemTimeInterruptMonitor example C# | Catch timeout exception when opening large Excel with Aspose.Cells | Abort Excel file load after 5 seconds
// Developer Intent: Configure an interrupt monitor to stop workbook loading after five seconds and handle the resulting exception.
// Use Cases: Prevent UI freeze in desktop applications by limiting workbook load time. | Enforce per‑file processing SLA in web services or APIs. | Log and skip files that exceed loading time during batch processing. | Provide immediate user feedback when a large Excel file takes too long to open.
// AI Prompts: Generate C# code that uses SystemTimeInterruptMonitor to stop loading a workbook after a configurable timeout and writes the exception to a log file. | Explain how to combine InterruptMonitor with async/await to load workbooks without blocking the main thread. | Show how to retry workbook loading with a larger timeout after an interruption occurs. | Provide guidance on customizing the interrupt monitor to abort based on memory usage instead of time.

using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Demonstrates how to create a SystemTimeInterruptMonitor, attach it to LoadOptions, start a 5‑second timer, and load a workbook. If loading exceeds the limit, the monitor throws an exception that is caught and handled gracefully.
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that will throw an exception when interrupted
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

            // Assign the monitor to LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring with a 5‑second (5000 ms) time limit
            monitor.StartMonitor(5000);

            try
            {
                // Attempt to load a workbook; the monitor will abort if loading exceeds 5 seconds
                Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

                // If loading succeeds, you can continue processing the workbook here
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Handle the interruption gracefully
                Console.WriteLine("Operation interrupted: " + ex.Message);
            }
        }
    }
}
