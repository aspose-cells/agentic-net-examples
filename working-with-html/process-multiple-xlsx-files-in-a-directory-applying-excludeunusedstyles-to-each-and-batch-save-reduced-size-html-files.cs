// Title: Batch convert XLSX to lightweight HTML with Aspose.Cells .NET (ExcludeUnusedStyles)
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, removes unused styles, sets HtmlSaveOptions.ExcludeUnusedStyles, and saves a matching .html file to an output directory while handling missing files and exceptions.
// Keywords: Aspose.Cells batch HTML export | ExcludeUnusedStyles .NET | remove unused styles Excel | convert multiple XLSX to HTML | reduce HTML size Aspose.Cells | C# Excel to HTML conversion | automated Excel HTML generation
// Common Searches: batch convert xlsx to html asp.net | aspocells excludeunusedstyles example | how to remove unused styles when saving html | c# export many excel files to html | optimize html size from excel workbook
// Developer Intent: Automatically transform every Excel workbook in a directory into a compact HTML file by stripping unused styles during the save process.
// Use Cases: Create web‑ready reports from a library of Excel templates in a single run. | Provide fast HTML previews for user‑uploaded spreadsheets in a portal, minimizing bandwidth. | Schedule nightly conversion of financial workbooks to lightweight HTML for quick visual review.
// AI Prompts: Generate a reusable C# method that processes all .xlsx files in a given folder and saves them as HTML with ExcludeUnusedStyles enabled. | Add robust logging and skip logic for corrupted or password‑protected Excel files in the batch conversion script. | Show how to preserve custom cell formatting while still excluding unused styles during HTML export with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, removes unused styles, sets HtmlSaveOptions.ExcludeUnusedStyles, and saves a matching .html file to an output directory while handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";

            // Output folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder '{sourceFolder}' does not exist.");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the source folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook from the current file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Optional: remove unused styles from the workbook to further reduce size
                    workbook.RemoveUnusedStyles();

                    // Configure HTML save options to exclude unused styles (default is true, set explicitly)
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExcludeUnusedStyles = true
                    };

                    // Build the output HTML file path (same file name with .html extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted '{xlsxPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
