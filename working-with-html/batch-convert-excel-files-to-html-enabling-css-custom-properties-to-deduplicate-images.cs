// Title: Batch convert .xlsx files to HTML with external image files using Aspose.Cells for .NET
// Common Searches: C# batch convert all .xlsx files in a folder to HTML using Aspose.Cells | How to export Excel images as separate files when saving to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example in C# | Convert multiple Excel workbooks to HTML without embedding images in .NET | Save Excel workbook as HTML with external image files using Aspose.Cells API
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | batch Excel to HTML conversion .NET | save workbook as HTML external images | convert folder of .xlsx to HTML C# | Aspose.Cells image export separate files

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to disable Base64 image embedding (ExportImagesAsBase64 = false), and saves each workbook as an HTML file with external image files into an output directory, handling missing files and exceptions.
class ExcelToHtmlBatchConverter
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcels";

        // Folder where the generated HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Process each .xlsx file in the input folder
        foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Ensure the Excel file still exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found, skipping: {excelPath}");
                    continue;
                }

                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Images will be saved as separate files (not Base64)
                    ExportImagesAsBase64 = false
                };

                // Determine the output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
