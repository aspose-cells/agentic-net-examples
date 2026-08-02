// Title: Abort a long‑running workbook load with ThreadInterruptMonitor in Aspose.Cells for .NET
// Description: Demonstrates how to create a ThreadInterruptMonitor, attach it to LoadOptions, start it with a timeout (e.g., 1000 ms), and load a large workbook. If the load exceeds the timeout, the monitor throws a CellsException with the Interrupted code, which can be caught to stop the operation; otherwise the monitor is stopped and the workbook is saved.
// Keywords: Aspose.Cells interrupt load | ThreadInterruptMonitor C# | LoadOptions InterruptMonitor example | cancel workbook loading .NET | CellsException Interrupted handling | timeout workbook load Aspose | abort Excel import Aspose.Cells
// Common Searches: how to stop workbook loading after timeout Aspose.Cells | ThreadInterruptMonitor usage with LoadOptions C# | catch CellsException.Interrupted during load | set time limit for loading large Excel file Aspose | cancel Excel import in .NET using Aspose.Cells
// Developer Intent: The developer needs to terminate a workbook loading process after a predefined time limit or user cancellation using Aspose.Cells' interrupt monitoring feature.
// Use Cases: Prevent UI freeze by limiting Excel file load to 1 second in a desktop app. | Allow users to cancel a long‑running import operation in a WPF or WinForms application. | Enforce per‑file processing timeouts in a server‑side batch service that handles many workbooks.
// AI Prompts: Provide a C# snippet that uses ThreadInterruptMonitor to abort a workbook load after 2 seconds and writes the interruption details to a log file. | Show how to wrap Workbook loading in a try‑catch block for CellsException.Interrupted and retry with different LoadOptions. | Explain how to combine a CancellationToken with Aspose.Cells InterruptMonitor for responsive UI cancellation in a .NET application.

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Demonstrates how to create a ThreadInterruptMonitor, attach it to LoadOptions, start it with a timeout (e.g., 1000 ms), and load a large workbook. If the load exceeds the timeout, the monitor throws a CellsException with the Interrupted code, which can be caught to stop the operation; otherwise the monitor is stopped and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that will throw an exception when interrupted
            ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);

            // Configure load options to use the monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start the monitor with a time limit (e.g., 1000 ms)
            monitor.StartMonitor(1000);

            try
            {
                // Attempt to load a workbook; this operation will be monitored
                Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

                // If loading completes before the time limit, stop the monitor
                monitor.FinishMonitor();

                // Optionally save the workbook (demonstrates normal flow)
                workbook.Save("LoadedResult.xlsx");
                Console.WriteLine("Workbook loaded and saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // The load operation was interrupted by the monitor
                Console.WriteLine("Loading operation was interrupted as expected.");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
