// Title: Load Excel with Aspose.Cells, disable charts via LoadOptions, set SystemTimeInterruptMonitor, and save as PDF (C#)
// Description: C# example that creates a SystemTimeInterruptMonitor, applies it to LoadOptions, loads an .xlsx file, removes all worksheet charts, and saves the workbook to PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | LoadOptions | disable charts | SystemTimeInterruptMonitor | interrupt monitor | PDF export | Excel to PDF | chart removal | performance optimization
// Common Searches: Aspose.Cells load workbook without charts | How to use SystemTimeInterruptMonitor in Aspose.Cells | Export Excel to PDF after removing charts C# | Set timeout for loading Excel with Aspose.Cells | Disable chart rendering Aspose.Cells .NET
// Developer Intent: Load an Excel file, suppress chart rendering, enforce a time‑out, and convert it to PDF using Aspose.Cells for .NET.
// Use Cases: Accelerate processing of large spreadsheets by skipping chart rendering. | Prevent long‑running load or save operations in automated batch jobs. | Generate PDF reports from workbooks where chart visuals are unnecessary. | Add timeout handling to Excel‑to‑PDF conversion pipelines.
// AI Prompts: Provide C# code that uses Aspose.Cells LoadOptions with a SystemTimeInterruptMonitor to load an .xlsx file, clear all charts, and save the workbook as PDF. | Explain how to configure a SystemTimeInterruptMonitor for both loading and saving in Aspose.Cells, including the exception thrown when the timeout expires. | Step‑by‑step guide to improve performance by disabling charts during workbook load and then exporting the result to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates loading a workbook with LoadOptions that disables charts,
    // configures an interrupt monitor, and saves the workbook as PDF.
    // C# example that creates a SystemTimeInterruptMonitor, applies it to LoadOptions, loads an .xlsx file, removes all worksheet charts, and saves the workbook to PDF using Aspose.Cells for .NET.
    public class LoadDisableChartsAndSavePdfDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // -----------------------------------------------------------------
            // 1. Create a SystemTimeInterruptMonitor.
            //    The monitor will be used for both loading and saving operations.
            // -----------------------------------------------------------------
            // terminateWithoutException = false -> an exception will be thrown
            // when the operation is interrupted.
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

            // Optional: start the monitor with a time limit (e.g., 10 seconds).
            // Adjust the timeout as needed.
            monitor.StartMonitor(10_000); // 10,000 ms = 10 seconds

            try
            {
                // -----------------------------------------------------------------
                // 2. Configure LoadOptions.
                //    Assign the interrupt monitor to the LoadOptions instance.
                // -----------------------------------------------------------------
                LoadOptions loadOptions = new LoadOptions
                {
                    InterruptMonitor = monitor
                };

                // -----------------------------------------------------------------
                // 3. Load the workbook using the constructor that accepts a file path
                //    and LoadOptions.
                // -----------------------------------------------------------------
                const string inputPath = "input.xlsx";
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                Workbook workbook = new Workbook(inputPath, loadOptions);

                // -----------------------------------------------------------------
                // 4. Disable all charts in the workbook.
                //    Iterate through each worksheet and clear its Charts collection.
                // -----------------------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Charts.Clear();
                }

                // -----------------------------------------------------------------
                // 5. Assign the same interrupt monitor to the workbook for the save
                //    operation. This allows the save to be interrupted if needed.
                // -----------------------------------------------------------------
                workbook.InterruptMonitor = monitor;

                // -----------------------------------------------------------------
                // 6. Save the workbook as PDF.
                // -----------------------------------------------------------------
                const string outputPath = "output.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine("Workbook loaded, charts removed, and saved to PDF successfully.");
            }
            finally
            {
                // -----------------------------------------------------------------
                // 7. Clean up.
                // -----------------------------------------------------------------
                // No explicit StopMonitor method in this version; the monitor will be
                // disposed when the application ends.
            }
        }
    }
}
