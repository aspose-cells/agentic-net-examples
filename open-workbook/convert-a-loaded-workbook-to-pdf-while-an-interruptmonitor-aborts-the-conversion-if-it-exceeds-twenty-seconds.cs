// Title: Save Excel to PDF with a 20‑second timeout using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx file, attach a ThreadInterruptMonitor that aborts the Save operation after 20 seconds, export the workbook to PDF, catch the interrupted CellsException, and properly finish the monitor.
// Keywords: Aspose.Cells | C# | .NET | ThreadInterruptMonitor | PDF conversion timeout | Excel to PDF | interrupt monitor | cancel workbook save | 20 second limit | CellsException
// Common Searches: Aspose.Cells set timeout for workbook.Save | C# abort PDF export after 20 seconds | ThreadInterruptMonitor example | How to cancel Excel to PDF conversion in .NET | Catch interrupted exception Aspose.Cells | Limit Aspose.Cells PDF generation time | Save workbook as PDF with time limit
// Developer Intent: The developer wants to export a loaded workbook to PDF but stop the operation if it runs longer than 20 seconds.
// Use Cases: Prevent long‑running PDF conversions in a web API by applying a 20‑second interrupt monitor. | Provide a cancellable export button in a desktop app that automatically halts after a timeout. | Guard server resources in batch processing pipelines by aborting oversized workbook saves.
// AI Prompts: Add logging to record the elapsed time before the interrupt occurs. | Show how to read the timeout value from an appsettings.json file and configure ThreadInterruptMonitor dynamically. | Explain handling of other CellsException codes when using an interrupt monitor during workbook.Save.

using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Shows how to load an .xlsx file, attach a ThreadInterruptMonitor that aborts the Save operation after 20 seconds, export the workbook to PDF, catch the interrupted CellsException, and properly finish the monitor.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create an interrupt monitor that will throw an exception when interrupted
            ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);

            // Assign the monitor to the workbook
            workbook.InterruptMonitor = monitor;

            // Set a time limit of 20,000 milliseconds (20 seconds) for the save operation
            monitor.StartMonitor(20000);

            try
            {
                // Attempt to save the workbook as PDF; this will be aborted if it exceeds the time limit
                workbook.Save("output.pdf", SaveFormat.Pdf);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // The operation was interrupted because it exceeded the time limit
                Console.WriteLine("Save operation was interrupted after exceeding the 20‑second limit.");
            }
            finally
            {
                // Ensure the monitor is finished for this operation
                monitor.FinishMonitor();
            }
        }
    }
}
