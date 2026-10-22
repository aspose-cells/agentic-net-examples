// Title: Convert a folder of Excel .xls and .xlsx files to HTML with grid lines using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through a directory, loads each .xls or .xlsx workbook with Aspose.Cells, and saves it as an HTML file with ExportGridLines enabled and images embedded as base64. | Adapt the batch converter to create a dedicated subfolder for each workbook's HTML output while keeping the original file names. | Add robust error handling that logs the full file path and exception details for any conversion failure and then continues processing the remaining files.
// Common Searches: how to batch convert excel files to html with grid lines using aspose.cells c# | c# aspnet convert multiple .xlsx files to html preserving cell borders | aspnet core process all excel workbooks in a folder and export to html with exportgridlines | save excel workbook as html with embedded images base64 asp.net cells library
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | batch conversion of .xls and .xlsx to HTML C# | embed images as base64 during Excel to HTML export | export all worksheets to HTML using Aspose.Cells | process workbooks in a folder with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlBatch
{
    // // This C# console application scans a specified input directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output directory. HtmlSaveOptions are set with ExportGridLines = true, ExportImagesAsBase64 = true, and ExportActiveWorksheetOnly = false to display grid lines, embed images, and export all worksheets.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel files
            string sourceFolder = @"C:\InputExcelFiles";

            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder '{sourceFolder}' does not exist.");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (supports .xls and .xlsx)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non‑Excel files

                // Verify the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Configure HTML save options to export grid lines
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = true,          // Show spreadsheet grid lines in HTML
                        ExportImagesAsBase64 = true,    // Embed images directly
                        ExportActiveWorksheetOnly = false // Export all worksheets
                    };

                    // Build the output HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
