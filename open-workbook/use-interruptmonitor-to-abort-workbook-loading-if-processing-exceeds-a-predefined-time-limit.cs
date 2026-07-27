// Title: Abort workbook loading with SystemTimeInterruptMonitor timeout in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 2‑second timer, and abort the Workbook constructor when loading a large Excel file exceeds the limit. The sample catches the resulting exception and shows optional post‑load handling.
// Keywords: Aspose.Cells | SystemTimeInterruptMonitor | interrupt monitor | load options timeout | abort workbook load | C# Excel timeout | large workbook performance | exception handling Aspose.Cells | Excel file load limit
// Common Searches: Aspose.Cells interrupt monitor example C# | How to stop workbook loading after timeout | SystemTimeInterruptMonitor usage with LoadOptions | Cancel large Excel file load in Aspose.Cells | Exception thrown by InterruptMonitor in Aspose.Cells
// Developer Intent: The developer needs to terminate the workbook loading process automatically when it exceeds a predefined time threshold.
// Use Cases: Prevent server‑side request timeouts by limiting Excel import duration in web APIs. | Provide responsive UI in desktop apps by aborting long‑running file opens. | Implement user‑initiated cancelation by linking a monitor to a Cancel button or CancellationToken.
// AI Prompts: Show how to replace SystemTimeInterruptMonitor with a custom monitor that checks a CancellationToken. | Write a try‑catch block that distinguishes a timeout interruption from other load errors. | Explain how to retrieve the elapsed time from SystemTimeInterruptMonitor after a successful load.

using System;
using Aspose.Cells;

// Demonstrates how to attach a SystemTimeInterruptMonitor to LoadOptions, start a 2‑second timer, and abort the Workbook constructor when loading a large Excel file exceeds the limit. The sample catches the resulting exception and shows optional post‑load handling.
class InterruptMonitorDemo
{
    static void Main()
    {
        // Create an interrupt monitor; false means an exception will be thrown on interruption
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Attach the monitor to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.InterruptMonitor = monitor;

        // Start monitoring with a time limit (e.g., 2000 ms = 2 seconds)
        monitor.StartMonitor(2000);

        try
        {
            // Load the workbook using the LoadOptions that contain the monitor
            Workbook wb = new Workbook("Large.xlsx", loadOptions);

            // If loading completes within the time limit, execution continues here
            Console.WriteLine("Workbook loaded successfully.");

            // Example save (optional)
            wb.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            // Loading was interrupted because it exceeded the time limit
            Console.WriteLine("Loading aborted: " + ex.Message);
        }
    }
}
