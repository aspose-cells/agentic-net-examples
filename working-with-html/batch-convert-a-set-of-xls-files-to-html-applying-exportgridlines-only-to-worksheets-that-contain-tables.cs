// Title: Batch convert multiple XLS workbooks to HTML with grid lines only on worksheets that contain tables using Aspose.Cells for .NET
// Common Searches: c# aspnet convert all xls files in a folder to html with aspose.cells | how to export gridlines only for sheets that have tables when saving workbook as html using aspose.cells | batch processing excel to html with conditional gridline visibility in c# | aspose.cells htmlsaveoptions exportimagesasbase64 example c#
// Tags: batch xls to html conversion aspose.cells | conditional worksheet gridlines aspose.cells | detect tables listobjects aspose.cells | htmlsaveoptions embed images base64 c# | c# iterate workbooks save as html aspose

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsToHtml
{
    // The C# program scans a specified input directory for .xls files, loads each workbook with Aspose.Cells, and configures HtmlSaveOptions to embed images as Base64. It iterates through every worksheet, checks the ListObjects collection to determine if a table exists, and sets IsGridlinesVisible accordingly—enabling grid lines only on sheets that contain tables. Each workbook is then saved as an HTML file with the same base name in an output folder, with comprehensive error handling and console logging for batch processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLS files
            string sourceFolder = @"C:\InputXls";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xls files in the source folder (including .xlsx if needed)
            string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

            foreach (string xlsPath in xlsFiles)
            {
                // Guard against missing files (should not happen with GetFiles, but added for safety)
                if (!File.Exists(xlsPath))
                {
                    Console.WriteLine($"File not found: {xlsPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsPath);

                    // Prepare HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Export grid lines globally; per‑worksheet visibility will be controlled via IsGridlinesVisible
                        ExportGridLines = true,
                        // Embed images as base64 to avoid external files
                        ExportImagesAsBase64 = true
                        // Note: SingleFile option is not available in this version of Aspose.Cells
                    };

                    // Iterate through each worksheet to set grid line visibility based on presence of tables
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // A worksheet contains tables if its ListObjects collection has any items
                        bool hasTable = sheet.ListObjects.Count > 0;

                        // Enable grid lines only for worksheets that contain tables
                        sheet.IsGridlinesVisible = hasTable;
                    }

                    // Determine output HTML file name (same base name as the XLS file)
                    string htmlFileName = Path.GetFileNameWithoutExtension(xlsPath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);
                    Console.WriteLine($"Converted: {xlsPath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
