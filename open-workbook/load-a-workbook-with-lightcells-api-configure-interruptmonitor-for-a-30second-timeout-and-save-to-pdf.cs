// Title: C# – Load Excel with LightCells API, apply a 30‑second SystemTimeInterruptMonitor, and export to PDF using Aspose.Cells
// Description: Shows how to set up a SystemTimeInterruptMonitor (termination disabled) with a 30‑second limit, attach it to LoadOptions, load an Excel workbook via the LightCells API, and save the workbook as a PDF file in .NET.
// Keywords: Aspose.Cells | LightCells API | SystemTimeInterruptMonitor | 30 second timeout | C# .NET | LoadOptions | Excel to PDF conversion | interrupt monitor example | prevent long load | GitHub sample
// Common Searches: Aspose.Cells LightCells interrupt monitor C# | set timeout for workbook load Aspose.Cells .NET | convert Excel to PDF with timeout | cancel long running load Aspose.Cells | SystemTimeInterruptMonitor usage example | load large Excel file with 30 sec limit | Aspose.Cells PDF export with interrupt monitor
// Developer Intent: Load an Excel workbook using LightCells with a 30‑second interrupt monitor and save it as a PDF.
// Use Cases: Avoid hanging when processing massive spreadsheets by enforcing a 30‑second load limit before conversion. | Implement time‑bounded Excel‑to‑PDF conversion in web APIs or micro‑services. | Gracefully cancel a long‑running load operation without throwing exceptions while still allowing the file to be saved.
// AI Prompts: Generate a C# snippet that stops a LightCells load after 20 seconds and logs the interruption before exporting to PDF. | Explain how to reuse a SystemTimeInterruptMonitor for both loading and saving steps in Aspose.Cells, including error handling. | Provide a step‑by‑step guide for configuring an interrupt monitor that terminates without exception and integrates with a CI/CD pipeline.

using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsInterruptDemo
{
    // Shows how to set up a SystemTimeInterruptMonitor (termination disabled) with a 30‑second limit, attach it to LoadOptions, load an Excel workbook via the LightCells API, and save the workbook as a PDF file in .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Create a SystemTimeInterruptMonitor with terminateWithoutException = false
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

            // Configure LoadOptions to use the interrupt monitor
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.InterruptMonitor = monitor;

            // Start the monitor with a 30‑second (30000 ms) time limit
            monitor.StartMonitor(30000);

            // Load the workbook using LightCells API (via LoadOptions)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Optionally restart the monitor before the save operation
            monitor.StartMonitor(30000);

            // Save the workbook to PDF format
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
