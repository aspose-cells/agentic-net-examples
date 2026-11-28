// Title: C# handling missing hidden worksheets when exporting to HTML with Aspose.Cells ExportHiddenWorksheet option
// Common Searches: Aspose.Cells C# export hidden worksheets to HTML and handle missing sheet exception | How to catch and retry HTML export when a hidden worksheet is not found in Aspose.Cells | C# code to disable ExportHiddenWorksheet after exception during workbook.Save to HTML | Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet error handling example | Detect missing hidden worksheet during HTML conversion with Aspose.Cells .NET
// Tags: ExportHiddenWorksheet exception handling | Aspose.Cells HTML export hidden sheet | HtmlSaveOptions retry missing worksheet | C# workbook.Save hidden worksheet fallback | detect missing hidden worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportHiddenWorksheetDemo
{
    // The sample loads a workbook, hides a sheet, removes another to simulate a missing hidden worksheet, sets HtmlSaveOptions with ExportHiddenWorksheet enabled, attempts to save to HTML, catches exceptions, identifies errors related to missing hidden worksheets, disables ExportHiddenWorksheet, retries the export, and logs the results while also checking file existence and handling unexpected errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (ensure the file exists in the execution folder)
            string sourceFile = "SampleWorkbook.xlsx";
            // Path for the generated HTML output
            string outputFile = "ExportedWorkbook.html";

            // Verify source file existence
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: Source file \"{sourceFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(sourceFile);

                // ------------------------------------------------------------
                // Simulate a scenario where a hidden worksheet might be missing.
                // ------------------------------------------------------------
                // Hide the second worksheet (if it exists)
                if (workbook.Worksheets.Count > 1)
                {
                    Worksheet hiddenSheet = workbook.Worksheets[1];
                    hiddenSheet.IsVisible = false;
                    Console.WriteLine($"Worksheet \"{hiddenSheet.Name}\" is now hidden.");
                }

                // Remove a worksheet that was previously hidden to create a "missing hidden worksheet" situation.
                // This mimics a user deleting a hidden sheet after ExportHiddenWorksheet was set to true.
                if (workbook.Worksheets.Count > 2)
                {
                    Worksheet toRemove = workbook.Worksheets[2];
                    Console.WriteLine($"Removing worksheet \"{toRemove.Name}\" (it may have been hidden).");
                    workbook.Worksheets.RemoveAt(2);
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Enable exporting of hidden worksheets
                    ExportHiddenWorksheet = true,
                    // Export the whole workbook (not only the active sheet)
                    ExportActiveWorksheetOnly = false
                };

                // Attempt to save with ExportHiddenWorksheet = true
                Console.WriteLine("Attempting to export workbook with hidden worksheets included...");
                try
                {
                    workbook.Save(outputFile, htmlOptions);
                    Console.WriteLine($"Export succeeded. HTML file saved to \"{outputFile}\".");
                }
                catch (Exception ex)
                {
                    // Specific handling for missing worksheet scenario
                    Console.WriteLine($"Export failed: {ex.Message}");
                    Console.WriteLine("Analyzing exception to determine if it is due to a missing hidden worksheet...");

                    if (ex.Message.Contains("hidden", StringComparison.OrdinalIgnoreCase) ||
                        ex.Message.Contains("worksheet", StringComparison.OrdinalIgnoreCase) ||
                        ex.Message.Contains("not found", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("Detected missing hidden worksheet. Adjusting options to exclude hidden worksheets and retrying.");

                        // Disable exporting hidden worksheets and retry
                        htmlOptions.ExportHiddenWorksheet = false;
                        try
                        {
                            workbook.Save(outputFile, htmlOptions);
                            Console.WriteLine($"Export succeeded after disabling hidden worksheet export. HTML file saved to \"{outputFile}\".");
                        }
                        catch (Exception retryEx)
                        {
                            Console.WriteLine($"Retry export also failed: {retryEx.Message}");
                        }
                    }
                    else
                    {
                        // Re‑throw if the exception is unrelated
                        Console.WriteLine("The error is unrelated to hidden worksheets. No further action taken.");
                    }
                }
            }
            catch (Exception outerEx)
            {
                // Catch any unexpected errors (e.g., file I/O, licensing)
                Console.WriteLine($"Unexpected error: {outerEx.Message}");
            }
        }
    }
}
