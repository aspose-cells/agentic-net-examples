// Title: Batch convert Excel workbooks to HTML with conditional ExportHiddenWorksheet for confidential sheets – Aspose.Cells C#
// Description: C# sample that loops through multiple Excel files, detects worksheets whose name contains "Confidential", sets HtmlSaveOptions.ExportHiddenWorksheet = false only for those workbooks, and saves each file as HTML using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | batch Excel to HTML | ExportHiddenWorksheet false | conditional HTML export | detect confidential worksheet | process multiple workbooks | HTMLSaveOptions | Aspose.Cells example | Excel hidden sheets
// Common Searches: Aspose.Cells batch export Excel to HTML | how to hide confidential worksheets when saving as HTML | C# set ExportHiddenWorksheet false conditionally | convert multiple workbooks to HTML with Aspose.Cells | detect worksheet name contains Confidential in C#
// Developer Intent: Automatically generate HTML previews for a collection of Excel files while ensuring that hidden worksheets are omitted only when a workbook contains a sheet labeled as confidential.
// Use Cases: Publish financial dashboards as HTML but suppress hidden confidential tabs. | Run nightly automation that converts uploaded reports to web‑ready pages, tightening export rules for sensitive workbooks. | Provide end‑users with HTML previews of their Excel uploads, preserving hidden sheets unless the file includes confidential data.
// AI Prompts: Create an async version of the batch processor that loads and saves workbooks concurrently while keeping the ExportHiddenWorksheet condition. | Add logging to record which files had confidential worksheets and were saved with hidden sheets excluded. | Extend the example to support other confidentiality indicators such as a custom property or cell value.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // C# sample that loops through multiple Excel files, detects worksheets whose name contains "Confidential", sets HtmlSaveOptions.ExportHiddenWorksheet = false only for those workbooks, and saves each file as HTML using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Example input: array of workbook file paths to process
            string[] workbookFiles = new string[]
            {
                @"C:\Workbooks\Report1.xlsx",
                @"C:\Workbooks\ConfidentialReport.xlsx",
                @"C:\Workbooks\Summary.xls"
            };

            // Output folder where processed HTML files will be saved
            string outputFolder = @"C:\ProcessedHtml";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each workbook
            foreach (string filePath in workbookFiles)
            {
                ProcessWorkbook(filePath, outputFolder);
            }

            Console.WriteLine("Batch processing completed.");
        }

        /// <param name="inputPath">Full path to the source workbook.</param>
        /// <param name="outputDir">Directory where the HTML output will be stored.</param>
        private static void ProcessWorkbook(string inputPath, string outputDir)
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the file (lifecycle: load)
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Determine if any worksheet name contains the word "Confidential"
                    bool hasConfidentialSheet = false;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            hasConfidentialSheet = true;
                            break;
                        }
                    }

                    // Prepare HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Apply ExportHiddenWorksheet = false only when a confidential sheet exists
                    if (hasConfidentialSheet)
                    {
                        saveOptions.ExportHiddenWorksheet = false;
                    }
                    // Otherwise keep the default value (true) – no need to set explicitly

                    // Build the output file name (same base name with .html extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(inputPath) + ".html";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    // Save the workbook as HTML using the configured options (lifecycle: save)
                    workbook.Save(outputPath, saveOptions);
                }
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
            }
        }
    }
}
