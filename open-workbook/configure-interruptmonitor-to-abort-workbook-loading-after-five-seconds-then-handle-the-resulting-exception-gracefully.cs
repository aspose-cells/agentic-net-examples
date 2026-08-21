// Title: Abort Aspose.Cells workbook load after 5 seconds with InterruptMonitor (C#)
// Description: Shows how to set up SystemTimeInterruptMonitor in Aspose.Cells for .NET, apply it to LoadOptions, enforce a 5‑second load timeout, and gracefully catch the interruption exception.
// Keywords: Aspose.Cells | InterruptMonitor | SystemTimeInterruptMonitor | C# | .NET | load timeout | abort workbook load | LoadOptions | exception handling | large Excel file
// Common Searches: Aspose.Cells stop workbook loading after timeout | SystemTimeInterruptMonitor example C# | How to cancel Excel load with Aspose.Cells | Catch timeout exception when loading workbook Aspose | Set load time limit for Aspose.Cells workbook
// Developer Intent: Configure an InterruptMonitor to abort workbook loading after five seconds and handle the resulting exception gracefully.
// Use Cases: Prevent UI freeze by limiting Excel load time in a desktop app. | Log a timeout event and switch to an alternative processing path when a workbook takes too long to load. | Enforce per‑file execution limits in a batch service that processes many workbooks.
// AI Prompts: Generate C# code that aborts workbook loading after 3 seconds using SystemTimeInterruptMonitor and logs the exception details. | Provide a retry pattern with exponential back‑off for a workbook load that timed out via InterruptMonitor. | Explain how to distinguish an InterruptMonitor timeout from other Aspose.Cells load errors in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Shows how to set up SystemTimeInterruptMonitor in Aspose.Cells for .NET, apply it to LoadOptions, enforce a 5‑second load timeout, and gracefully catch the interruption exception.
    class Program
    {
        static void Main()
        {
            // Create an interrupt monitor that will throw an exception when interrupted
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

            // Assign the monitor to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.InterruptMonitor = monitor;

            // Start monitoring with a 5‑second (5000 ms) time limit
            monitor.StartMonitor(5000);

            try
            {
                // Attempt to load a workbook; the monitor will abort if loading exceeds 5 seconds
                Workbook workbook = new Workbook("sample.xlsx", loadOptions);

                // If loading succeeds, you can continue processing the workbook here
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Handle the interruption (or any other exception) gracefully
                Console.WriteLine("Operation interrupted: " + ex.Message);
            }
        }
    }
}
