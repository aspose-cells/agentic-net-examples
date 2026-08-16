// Title: C# – Load Excel with LightCells API, set a 30‑second SystemTimeInterruptMonitor, and save as PDF using Aspose.Cells
// Description: Demonstrates how to configure a SystemTimeInterruptMonitor for a 30‑second timeout, attach it to LoadOptions, load an Excel workbook via the LightCells API, and export the workbook to PDF with Aspose.Cells for .NET. Ideal for server‑side scenarios where long‑running loads must be bounded.
// Keywords: Aspose.Cells | LightCells API | SystemTimeInterruptMonitor | LoadOptions | C# | Excel to PDF conversion | timeout monitoring | workbook load timeout | server side Excel processing | .NET PDF export
// Common Searches: Aspose.Cells LightCells interrupt monitor example | C# set timeout when loading Excel workbook | Convert Excel to PDF with LightCells and timeout | SystemTimeInterruptMonitor 30 seconds Aspose | LoadOptions InterruptMonitor usage C#
// Developer Intent: Load an Excel file with LightCells while enforcing a 30‑second timeout, then convert the workbook to PDF.
// Use Cases: Prevent runaway workbook loads in web services or background jobs. | Enforce execution‑time limits for user‑uploaded Excel files in multi‑tenant SaaS platforms. | Guarantee PDF conversion completes within a predefined window for batch processing pipelines.
// AI Prompts: Show how to catch the timeout exception thrown by SystemTimeInterruptMonitor and return a custom error response. | Generate logging code that records when the interrupt monitor aborts a workbook load. | Explain how to replace the time‑based monitor with a CPU‑usage based InterruptMonitor in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to configure a SystemTimeInterruptMonitor for a 30‑second timeout, attach it to LoadOptions, load an Excel workbook via the LightCells API, and export the workbook to PDF with Aspose.Cells for .NET. Ideal for server‑side scenarios where long‑running loads must be bounded.
class LightCellsInterruptDemo
{
    static void Main()
    {
        // Create a SystemTimeInterruptMonitor (false = do not terminate without exception)
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Configure LoadOptions to use the interrupt monitor
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start the monitor with a 30‑second (30000 ms) time limit
        monitor.StartMonitor(30000);

        // Load the workbook using the LightCells API (enabled via LoadOptions)
        // Replace "input.xlsx" with the path to your source Excel file
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook as PDF
        // Replace "output.pdf" with the desired output path
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
