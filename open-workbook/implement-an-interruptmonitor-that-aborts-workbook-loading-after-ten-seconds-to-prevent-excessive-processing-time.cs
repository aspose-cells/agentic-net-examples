// Title: Abort Excel workbook loading after 10 seconds with Aspose.Cells InterruptMonitor (C#)
// Description: Shows how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 10‑second timer, and load a workbook. If loading exceeds the limit, the monitor throws an exception that is caught, preventing excessive processing time.
// Keywords: Aspose.Cells | InterruptMonitor | SystemTimeInterruptMonitor | C# timeout | cancel workbook load | LoadOptions | Excel file loading | time‑limited load | exception handling | large Excel performance
// Common Searches: Aspose.Cells set timeout for workbook load | C# interrupt monitor abort Excel loading | How to stop loading large Excel file after 10 seconds | LoadOptions InterruptMonitor example | Cancel workbook loading with Aspose.Cells
// Developer Intent: Implement a 10‑second timeout that aborts workbook loading.
// Use Cases: Web services that must reject oversized Excel uploads after a short processing window. | Batch import jobs that need to enforce strict execution time limits. | User‑driven cancel operation for long‑running workbook loads in desktop applications.
// AI Prompts: Generate C# code for a custom InterruptMonitor that logs the interruption before throwing. | Provide a pattern to retry workbook loading with LoadDataOnly after a timeout occurs. | Explain how to combine InterruptMonitor with asynchronous loading in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 10‑second timer, and load a workbook. If loading exceeds the limit, the monitor throws an exception that is caught, preventing excessive processing time.
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

        // Start monitoring with a 10‑second (10000 ms) time limit
        monitor.StartMonitor(10000);

        try
        {
            // Load the workbook using the load options that contain the monitor
            Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);

            // If loading completes within the time limit, optionally save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            // Loading was interrupted after exceeding the time limit
            Console.WriteLine("Loading aborted: " + ex.Message);
        }
    }
}
