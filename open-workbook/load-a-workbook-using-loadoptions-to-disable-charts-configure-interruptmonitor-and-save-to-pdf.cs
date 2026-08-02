// Title: Load Excel Workbook without Charts, Apply InterruptMonitor, and Save as PDF – Aspose.Cells for .NET
// Description: Demonstrates how to create a LoadOptions object that skips chart rendering, attach a SystemTimeInterruptMonitor with a timeout, load an Excel file, and export it to PDF while handling interruption and other exceptions in C#.
// Keywords: Aspose.Cells LoadOptions disable charts | SystemTimeInterruptMonitor C# | Excel to PDF conversion .NET | interrupt monitor timeout Aspose.Cells | load workbook without charts | exception handling CellsException | C# Aspose.Cells PDF export | prevent long-running Excel processing
// Common Searches: Aspose.Cells load workbook without charts | How to use InterruptMonitor with Aspose.Cells | Export Excel to PDF while ignoring charts .NET | Set timeout for Excel processing Aspose.Cells | C# example for LoadOptions and PDF save
// Developer Intent: Load an Excel file with charts suppressed, enforce a processing timeout, and convert it to PDF.
// Use Cases: Generate PDF reports from large workbooks without the overhead of chart rendering. | Abort lengthy load or save operations after a predefined time to keep services responsive. | Provide graceful error handling for interrupted Excel processing in enterprise applications.
// AI Prompts: Create C# code that loads an Excel workbook using LoadOptions to skip charts, applies a SystemTimeInterruptMonitor with a 5‑second limit, and saves the workbook as a PDF. | Show how to catch CellsException when an Aspose.Cells operation is interrupted by an InterruptMonitor. | Explain how to configure LoadOptions for chart suppression and reuse the same InterruptMonitor for both loading and saving steps.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsDemo
{
    // Demonstrates how to create a LoadOptions object that skips chart rendering, attach a SystemTimeInterruptMonitor with a timeout, load an Excel file, and export it to PDF while handling interruption and other exceptions in C#.
    class LoadDisableChartsAndSavePdf
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Create LoadOptions (no LoadDataOnly property in this version)
            LoadOptions loadOptions = new LoadOptions();

            // Set up an interrupt monitor to allow operation cancellation
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
            loadOptions.InterruptMonitor = monitor;

            try
            {
                // Start the monitor with a time limit (e.g., 5 seconds)
                monitor.StartMonitor(5000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start interrupt monitor: {ex.Message}");
                // Continue without monitor if it fails
            }

            try
            {
                // Load the workbook using the configured LoadOptions
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Assign the same interrupt monitor to the workbook (optional but ensures
                // the monitor is also used during save operations)
                workbook.InterruptMonitor = monitor;

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine("Workbook loaded (charts disabled) and saved to PDF successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Handle operation interruption
                Console.WriteLine("The operation was interrupted by the monitor.");
            }
            catch (Exception ex)
            {
                // Handle other possible exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
