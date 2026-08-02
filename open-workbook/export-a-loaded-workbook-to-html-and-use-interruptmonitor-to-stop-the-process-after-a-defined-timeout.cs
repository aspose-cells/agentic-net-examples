// Title: C# – Export Workbook to HTML with a 2‑second Timeout Using Aspose.Cells ThreadInterruptMonitor
// Description: Demonstrates how to load an Excel file and save it as HTML while enforcing a 2‑second execution limit. The sample configures a ThreadInterruptMonitor in LoadOptions, reuses it for the Workbook save operation, and catches timeout exceptions to keep the application responsive.
// Keywords: Aspose.Cells C# HTML export timeout | ThreadInterruptMonitor example | limit workbook load time Aspose.Cells | interrupt Excel to HTML conversion | C# save workbook with timeout
// Common Searches: Aspose.Cells export to HTML with timeout C# | ThreadInterruptMonitor load and save example | how to stop Aspose.Cells save after milliseconds | C# limit Excel conversion time Aspose.Cells | interrupt workbook loading Aspose.Cells
// Developer Intent: Load an Excel workbook and convert it to HTML while guaranteeing that neither the load nor the save exceeds a predefined time limit.
// Use Cases: Prevent long‑running conversions from freezing a UI or service. | Enforce processing budgets for large Excel files in batch jobs. | Gracefully handle timeout events and log or retry the operation.
// AI Prompts: Write C# code that uses Aspose.Cells ThreadInterruptMonitor to abort a workbook load or HTML export after a specific timeout. | Explain the steps to attach an InterruptMonitor to LoadOptions and to a Workbook for both loading and saving. | Refactor the example to use a using statement for the monitor and ensure all resources are disposed correctly.

using System;
using Aspose.Cells;

// Demonstrates how to load an Excel file and save it as HTML while enforcing a 2‑second execution limit. The sample configures a ThreadInterruptMonitor in LoadOptions, reuses it for the Workbook save operation, and catches timeout exceptions to keep the application responsive.
class ExportWorkbookToHtmlWithInterrupt
{
    static void Main()
    {
        // Timeout for each operation (in milliseconds)
        int timeoutMs = 2000; // 2 seconds

        // Create an interrupt monitor (set terminateWithoutException to false)
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);

        // Prepare load options with the monitor
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring before loading the workbook
        monitor.StartMonitor(timeoutMs);
        Workbook workbook;
        try
        {
            // Load the workbook (monitor will interrupt if loading exceeds timeout)
            workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);
            monitor.FinishMonitor(); // Clean up monitor thread after loading
        }
        catch (Exception ex)
        {
            Console.WriteLine("Loading was interrupted: " + ex.Message);
            return;
        }

        // Assign the same monitor to the workbook for the save operation
        workbook.InterruptMonitor = monitor;

        // Start monitoring before saving to HTML
        monitor.StartMonitor(timeoutMs);
        try
        {
            // Export the workbook to HTML (monitor will interrupt if saving exceeds timeout)
            workbook.Save("Result.html", SaveFormat.Html);
            monitor.FinishMonitor(); // Clean up monitor thread after saving
            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Saving was interrupted: " + ex.Message);
        }
    }
}
