// Title: Export only visible worksheets to separate HTML files using Aspose.Cells for .NET (ExportHiddenWorksheet = false)
// AI Prompts: Write C# code that creates a workbook with both visible and hidden worksheets, configures HtmlSaveOptions.ExportHiddenWorksheet to false, and saves each visible sheet as an individual HTML file. | Demonstrate how to programmatically verify that the output folder contains HTML files only for the visible worksheets after the export. | Modify the example to export hidden worksheets as well by setting ExportHiddenWorksheet to true and show the difference in the generated files.
// Common Searches: Aspose.Cells C# export workbook to HTML without hidden sheets | how to save each visible worksheet as a separate HTML file using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet false example | verify that hidden worksheets are excluded from HTML export in Aspose.Cells | generate per‑sheet HTML output while ignoring hidden worksheets in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet false | C# export workbook to separate HTML files per worksheet | skip hidden worksheets during HTML conversion Aspose.Cells | validate HTML output contains only visible sheets | per‑sheet HTML generation with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportExample
{
    // The example creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook to a folder (producing separate HTML files only for visible sheets), and then checks that the hidden worksheet was not exported.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ----- Visible worksheet -----
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Visible Content");

                // ----- Hidden worksheet -----
                int hiddenIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
                hiddenSheet.Name = "HiddenSheet";
                hiddenSheet.Cells["A1"].PutValue("Hidden Content");
                hiddenSheet.IsVisible = false; // Mark as hidden

                // Define HTML save options (do not export hidden sheets)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportHiddenWorksheet = false
                };

                // Choose an output folder and ensure it exists
                string outputFolder = @"C:\Temp\AsposeExportTest";
                Directory.CreateDirectory(outputFolder);

                // Save the workbook to HTML (one file per worksheet)
                // When a folder path is supplied, Aspose.Cells creates a separate HTML file for each visible sheet.
                workbook.Save(outputFolder, htmlOptions);

                // Verify that only the visible worksheet was exported
                string[] exportedFiles = Directory.GetFiles(outputFolder, "*.html");

                // Expect only one HTML file (the visible sheet)
                if (exportedFiles.Length == 1 &&
                    Path.GetFileNameWithoutExtension(exportedFiles[0])
                        .Equals("VisibleSheet", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Test Passed: Hidden worksheet was not exported.");
                }
                else
                {
                    Console.WriteLine("Test Failed: Hidden worksheet export behavior is incorrect.");
                }

                // Optional clean up
                // Directory.Delete(outputFolder, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
