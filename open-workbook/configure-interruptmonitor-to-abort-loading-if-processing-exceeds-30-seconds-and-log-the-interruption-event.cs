// Title: Abort Aspose.Cells workbook load after 30 seconds with SystemTimeInterruptMonitor (C#)
// Description: Shows how to attach a SystemTimeInterruptMonitor to LoadOptions, stop loading a large Excel file when it exceeds a 30‑second limit, and log the interruption exception.
// Keywords: Aspose.Cells | SystemTimeInterruptMonitor | InterruptMonitor | C# | timeout | abort load | LoadOptions | Excel workbook | performance | exception logging
// Common Searches: Aspose.Cells timeout loading workbook | SystemTimeInterruptMonitor example C# | cancel workbook load after time limit Aspose | InterruptMonitor abort load Aspose.Cells | log load interruption Aspose.Cells
// Developer Intent: Set up a time‑based interrupt monitor that cancels workbook loading after 30 seconds and captures the exception for logging.
// Use Cases: Limit execution time for loading huge Excel files in web or API services. | Enforce SLA thresholds during batch processing of multiple workbooks. | Create diagnostic logs whenever a load operation exceeds expected duration. | Combine with terminateWithoutException=false to raise an exception for custom error handling.
// AI Prompts: Generate C# code that uses SystemTimeInterruptMonitor with a 30‑second timeout and writes the exception details to a log file via Serilog. | Explain how to configure InterruptMonitor with terminateWithoutException=true and check IsInterrupted after loading. | Provide a retry pattern for workbook loading that respects the interrupt monitor and logs each timeout event.

using System;
using Aspose.Cells;

// Shows how to attach a SystemTimeInterruptMonitor to LoadOptions, stop loading a large Excel file when it exceeds a 30‑second limit, and log the interruption exception.
class Program
{
    static void Main()
    {
        // Create an interrupt monitor that throws an exception when interrupted
        var monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);

        // Assign the monitor to LoadOptions
        var loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring with a 30‑second (30000 ms) time limit
        monitor.StartMonitor(30000);

        try
        {
            // Load the workbook using the configured LoadOptions
            var workbook = new Workbook("Large.xlsx", loadOptions);

            // Optionally perform additional processing here

            // Save the workbook (no monitor needed for this example)
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            // Log the interruption event
            Console.WriteLine($"Loading aborted due to interruption: {ex.Message}");
        }
    }
}
