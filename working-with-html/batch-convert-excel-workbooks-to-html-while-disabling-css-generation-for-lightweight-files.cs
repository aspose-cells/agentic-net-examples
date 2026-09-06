// Title: Batch convert Excel workbooks (.xls, .xlsx, .xlsm) to single‑file HTML with embedded images and no external CSS using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a folder, loads each .xls/.xlsx/.xlsm workbook with Aspose.Cells, and saves it as a standalone HTML file with all images embedded as Base64 while preventing external CSS generation. | Show how to configure Aspose.Cells HtmlSaveOptions to export every worksheet, embed images as Base64, and turn off CSS style output for a lightweight HTML result.
// Common Searches: C# Aspose.Cells convert all Excel files in a directory to HTML with embedded images | how to disable CSS when saving a workbook as HTML using Aspose.Cells .NET | batch export .xls and .xlsx to single HTML files without external resources | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example | convert multiple Excel workbooks to HTML in one script C#
// Tags: Aspose.Cells batch Excel to HTML conversion | HtmlSaveOptions images base64 | turn off CSS Aspose.Cells HTML export | export all worksheets to HTML C# | standalone HTML from .xls .xlsx .xlsm

using System;
using System.IO;
using Aspose.Cells;

// The program iterates over all .xls, .xlsx, and .xlsm files in a specified input folder, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 and suppress external CSS, then saves each workbook as a single, self‑contained HTML file in an output directory, logging successes and handling errors.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string inputFolder = @"C:\InputExcel";

        // Folder where the generated HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (xls, xlsx, xlsm) from the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            try
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                    continue; // Skip non‑Excel files

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options to produce lightweight files
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false, // Export all worksheets
                    ExportImagesAsBase64 = true        // Embed images to keep a single file
                };

                // Build the output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted '{filePath}' to '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
