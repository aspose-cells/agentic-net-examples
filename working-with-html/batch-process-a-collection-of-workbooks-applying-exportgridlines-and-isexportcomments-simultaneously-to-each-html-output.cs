// Title: Batch Convert Excel Workbooks to HTML with Gridlines & Comments – Aspose.Cells for .NET
// Description: C# utility that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, enables HtmlSaveOptions.ExportGridLines and HtmlSaveOptions.IsExportComments, and saves the result as HTML. Includes folder validation and exception handling for robust batch processing.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | ExportGridLines | IsExportComments | batch conversion | multiple workbooks | HtmlSaveOptions example | GitHub code sample
// Common Searches: Aspose.Cells batch export to HTML with gridlines and comments | C# convert all .xlsx files in a directory to HTML | HtmlSaveOptions ExportGridLines IsExportComments together | How to export Excel comments when saving as HTML using Aspose.Cells | Sample code for bulk Excel to HTML conversion .NET
// Developer Intent: Convert every Excel file in a specified folder to an HTML page that retains both gridlines and cell comments.
// Use Cases: Publish a set of financial spreadsheets on an intranet with original layout and reviewer notes. | Automate nightly generation of HTML documentation from design specs stored in Excel. | Create a web‑accessible archive of uploaded Excel reports while preserving visual cues and annotations.
// AI Prompts: Generate a C# script using Aspose.Cells that batch converts .xlsx files to HTML with gridlines and comments, and logs missing files. | Show how to modify the example so each worksheet is saved as a separate HTML file while keeping gridlines and comments. | Explain how to add HtmlSaveOptions such as ExportImages, ExportActiveWorksheet, or custom CSS without disabling ExportGridLines or IsExportComments.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // C# utility that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, enables HtmlSaveOptions.ExportGridLines and HtmlSaveOptions.IsExportComments, and saves the result as HTML. Includes folder validation and exception handling for robust batch processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder that contains the source Excel files
            string sourceFolder = @"C:\InputWorkbooks";

            // Define the folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel files (you can adjust the search pattern as needed)
                List<string> workbookFiles = new List<string>(Directory.GetFiles(sourceFolder, "*.xlsx"));

                // Process each workbook
                foreach (string workbookPath in workbookFiles)
                {
                    try
                    {
                        // Verify the workbook file exists before loading
                        if (!File.Exists(workbookPath))
                        {
                            Console.WriteLine($"File not found: {workbookPath}");
                            continue;
                        }

                        // Load the workbook from file
                        Workbook workbook = new Workbook(workbookPath);

                        // Create HTML save options and enable both gridlines and comments export
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            ExportGridLines = true,      // Export the gridlines
                            IsExportComments = true      // Export cell comments
                        };

                        // Build the output HTML file name (same base name as the workbook)
                        string outputFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the workbook as HTML using the configured options
                        workbook.Save(outputPath, htmlOptions);

                        Console.WriteLine($"Saved HTML for '{workbookPath}' to '{outputPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
