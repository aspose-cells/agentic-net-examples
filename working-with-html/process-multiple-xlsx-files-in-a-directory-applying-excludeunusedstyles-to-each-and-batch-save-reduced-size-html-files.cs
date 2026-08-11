// Title: Batch Convert XLSX Files to Compact HTML with Aspose.Cells (C#) – Remove Unused Styles
// Description: A C# console utility that scans a folder for *.xlsx workbooks, loads each with Aspose.Cells, calls RemoveUnusedStyles, sets HtmlSaveOptions.ExcludeUnusedStyles, and saves a reduced‑size HTML file to a target directory. Includes progress logging and robust error handling.
// Keywords: Aspose.Cells | C# batch export | XLSX to HTML conversion | RemoveUnusedStyles | ExcludeUnusedStyles | HTML size reduction | .NET console application | directory processing | Excel workbook HTML export | automated Excel to HTML
// Common Searches: batch convert xlsx to html aspose.cells | remove unused styles when saving excel as html | htmlsaveoptions excludeunusedstyles c# example | process multiple excel files in a folder with aspose.cells | reduce html file size generated from excel
// Developer Intent: Automatically transform every XLSX workbook in a specified folder into an HTML file while stripping unused styles to minimize output size.
// Use Cases: Generate lightweight HTML reports from a repository of Excel templates for web publishing. | Schedule nightly export of financial spreadsheets to HTML, ensuring only essential styles are retained to save bandwidth. | Pre‑process user‑uploaded Excel files for preview in a web portal, delivering fast‑loading HTML previews.
// AI Prompts: Write C# code that iterates through a directory of .xlsx files, removes unused styles, and saves each workbook as HTML using Aspose.Cells with ExcludeUnusedStyles enabled. | Explain how HtmlSaveOptions.ExcludeUnusedStyles affects the size and rendering of the generated HTML and when it should be applied. | Suggest best‑practice error‑handling patterns for a batch Excel‑to‑HTML conversion tool built with Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that scans a folder for *.xlsx workbooks, loads each with Aspose.Cells, calls RemoveUnusedStyles, sets HtmlSaveOptions.ExcludeUnusedStyles, and saves a reduced‑size HTML file to a target directory. Includes progress logging and robust error handling.
class BatchHtmlExport
{
    static void Main()
    {
        // Directory containing the source XLSX files
        string inputDirectory = @"C:\InputXlsx";

        // Directory where the reduced‑size HTML files will be saved
        string outputDirectory = @"C:\OutputHtml";

        // Verify input directory exists
        if (!Directory.Exists(inputDirectory))
        {
            Console.WriteLine($"Input directory not found: {inputDirectory}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        try
        {
            // Retrieve all XLSX files in the input directory
            string[] xlsxFiles = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Verify the file still exists before processing
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found (skipped): {xlsxPath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the XLSX file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Remove any styles that are not used in the workbook
                    workbook.RemoveUnusedStyles();

                    // Configure HTML save options to exclude unused styles
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExcludeUnusedStyles = true // default is true, set for clarity
                    };

                    // Build the output HTML file path (same name as source, different extension)
                    string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
                    string htmlPath = Path.Combine(outputDirectory, htmlFileName);

                    // Save the workbook as an HTML file using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {htmlFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("All XLSX files have been processed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
