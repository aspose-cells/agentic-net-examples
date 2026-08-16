// Title: Abort workbook loading after 30 seconds with SystemTimeInterruptMonitor in Aspose.Cells for .NET
// Description: Demonstrates how to create a SystemTimeInterruptMonitor with a 30‑second timeout, attach it to LoadOptions, and load a large Excel file. If loading exceeds the limit, a CellsException (ExceptionType.Interrupted) is thrown, caught, and logged, while other exceptions are also handled.
// Keywords: Aspose.Cells interrupt monitor | SystemTimeInterruptMonitor timeout | C# workbook load timeout | CellsException Interrupted handling | log workbook load interruption | prevent long‑running Excel load
// Common Searches: Aspose.Cells stop loading after time limit | SystemTimeInterruptMonitor 30 seconds example | catch CellsException Interrupted C# | log timeout when loading large Excel file Aspose | how to abort workbook load in .NET
// Developer Intent: Implement a time‑based interrupt monitor that aborts workbook loading after 30 seconds and records the interruption event.
// Use Cases: Safely load very large spreadsheets with a predefined timeout to avoid hanging services. | Enforce execution time limits in web APIs that process user‑uploaded Excel files. | Capture and log timeout events for performance monitoring and diagnostics.
// AI Prompts: Show C# code using SystemTimeInterruptMonitor with a 30‑second limit to abort Aspose.Cells workbook loading and log the interruption. | Provide an example that catches CellsException where ExceptionType is Interrupted and writes the message to a log file. | Explain how to configure InterruptMonitor to terminate silently without throwing an exception and how the error‑handling code changes.

using System;
using Aspose.Cells;

namespace InterruptMonitorExample
{
    // Demonstrates how to create a SystemTimeInterruptMonitor with a 30‑second timeout, attach it to LoadOptions, and load a large Excel file. If loading exceeds the limit, a CellsException (ExceptionType.Interrupted) is thrown, caught, and logged, while other exceptions are also handled.
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that throws an exception when interrupted
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);

            // Configure load options to use the monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start monitoring with a 30‑second time limit (30000 ms)
            monitor.StartMonitor(30000);

            try
            {
                // Attempt to load a workbook; the monitor will abort if loading exceeds 30 seconds
                Workbook workbook = new Workbook("Large.xlsx", loadOptions);

                // If loading succeeds, optionally save the workbook (not required for the interrupt demonstration)
                workbook.Save("Result.xlsx");
                Console.WriteLine("Workbook loaded and saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Log the interruption event
                Console.WriteLine("Loading was interrupted due to time limit: " + ex.Message);
            }
            catch (Exception ex)
            {
                // Log any other unexpected exceptions
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
