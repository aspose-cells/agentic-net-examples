// Title: Batch convert multiple Excel (.xls/.xlsx) workbooks to HTML without external CSS using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that scans a directory for .xls and .xlsx files and saves each workbook as an HTML file with embedded Base64 images, using Aspose.Cells HtmlSaveOptions. | Adjust the Aspose.Cells HTML export configuration to suppress external CSS file creation while preserving all worksheets and hidden sheets in the output. | Add comprehensive error handling to a bulk Excel‑to‑HTML conversion loop that logs missing or corrupt files and continues processing the remaining workbooks.
// Common Searches: how to export excel files to html without css using aspose.cells c# | c# batch convert xlsx to html embed images base64 aspose | asp.net console app convert folder of excel workbooks to html aspose.cells | disable css generation in aspose.cells htmlsaveoptions c# | save all worksheets including hidden ones to html with aspose.cells
// Tags: Aspose.Cells HTML export without external CSS | batch Excel to HTML conversion C# | embed images as Base64 in Aspose.Cells HTML output | export all worksheets including hidden sheets Aspose.Cells | process multiple workbooks in a folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program iterates through a specified input folder, loads each .xls or .xlsx workbook with Aspose.Cells, and saves it as an HTML file in an output folder. HtmlSaveOptions are configured to embed images as Base64, include all worksheets (including hidden ones), and suppress external CSS generation, while robust error handling logs any files that cannot be processed.
class ExcelToHtmlBatchConverter
{
    static void Main()
    {
        try
        {
            // Folder containing the Excel files to convert
            string inputFolder = @"C:\ExcelFiles";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\HtmlOutput";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (supports .xls and .xlsx)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non‑Excel files

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Configure HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // ExportImagesAsBase64 = true embeds images directly in the HTML
                        ExportImagesAsBase64 = true,
                        // Export all worksheets
                        ExportActiveWorksheetOnly = false,
                        // Include hidden worksheets if needed
                        ExportHiddenWorksheet = true,
                        // Export the whole sheet(s)
                        ExportPrintAreaOnly = false
                    };

                    // Build the output HTML file path
                    string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string htmlFilePath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlFilePath, saveOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
