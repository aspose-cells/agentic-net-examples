// Title: Convert a folder of Excel (.xls/.xlsx) files to HTML with grid lines using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch convert Excel files in a directory to HTML with grid lines | How to export all worksheets as HTML with visible grid lines using Aspose.Cells | Save multiple .xlsx workbooks to HTML preserving spreadsheet grid lines in .NET | Aspose.Cells HtmlSaveOptions ExportGridLines example for bulk conversion | Convert every Excel file in a folder to HTML with grid lines using C#
// Tags: Aspose.Cells batch Excel to HTML conversion | HtmlSaveOptions ExportGridLines C# | convert folder of .xlsx to .html programmatically | export all worksheets as HTML with grid lines | embed images as Base64 Aspose.Cells HTML export

using System;
using System.IO;
using Aspose.Cells;

// The C# utility scans a specified input directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to enable ExportGridLines (and optionally embed images as Base64 and export all worksheets), and saves each workbook as an .html file in an output directory, handling missing files and runtime errors.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing the Excel files to convert
        string inputFolder = @"C:\InputExcelFiles";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtmlFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        try
        {
            // Get all Excel files (xls, xlsx) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non‑Excel files

                // Verify the file still exists before loading
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
                        ExportGridLines = true,               // Show spreadsheet grid lines in HTML
                        ExportImagesAsBase64 = true,          // Embed images directly (optional)
                        ExportActiveWorksheetOnly = false    // Export all worksheets (adjust as needed)
                    };

                    // Determine output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{filePath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
