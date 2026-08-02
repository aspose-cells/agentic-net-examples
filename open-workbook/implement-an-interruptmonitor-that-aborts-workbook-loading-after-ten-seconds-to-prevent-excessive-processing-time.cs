// Title: Abort Workbook Loading in Aspose.Cells after 10 seconds with SystemTimeInterruptMonitor (C#)
// Description: Demonstrates how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 10‑second timer, and abort workbook loading with an exception if the operation exceeds the limit. Includes proper try‑catch handling for graceful fallback.
// Keywords: Aspose.Cells | InterruptMonitor | SystemTimeInterruptMonitor | C# workbook timeout | load Excel file with time limit | cancel long running load | large Excel file performance | ASP.NET Excel processing | US developers | European .NET community
// Common Searches: Aspose.Cells set timeout for loading workbook | how to stop Excel file load after 10 seconds in C# | SystemTimeInterruptMonitor example | prevent long workbook load Aspose.Cells | interrupt monitor usage Aspose.Cells .NET
// Developer Intent: Create a 10‑second timeout that aborts workbook loading using Aspose.Cells' InterruptMonitor.
// Use Cases: Terminate excessively long loads of massive Excel files in a web service to keep response times low. | Enforce processing time caps in batch import jobs that read workbooks, avoiding server overload. | Provide immediate feedback to users when a file cannot be opened within the allowed period.
// AI Prompts: Generate code that logs the elapsed time before the SystemTimeInterruptMonitor aborts the load. | Show how to retry workbook loading with an increased timeout after an interruption. | Explain how to catch the specific Aspose.Cells.InterruptException and handle it gracefully.

using System;
using Aspose.Cells;

// Demonstrates how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 10‑second timer, and abort workbook loading with an exception if the operation exceeds the limit. Includes proper try‑catch handling for graceful fallback.
class Program
{
    static void Main()
    {
        // Create an interrupt monitor that will throw an exception when interrupted
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Attach the monitor to load options
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring with a 10‑second (10000 ms) time limit
        monitor.StartMonitor(10000);

        try
        {
            // Attempt to load a workbook using the load options with the monitor attached
            Workbook workbook = new Workbook("Large.xlsx", loadOptions);

            // If loading completes within the time limit, you can continue processing here
            // For demonstration, we simply indicate success
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Loading was interrupted after exceeding the time limit
            Console.WriteLine("Loading aborted: " + ex.Message);
        }
    }
}
