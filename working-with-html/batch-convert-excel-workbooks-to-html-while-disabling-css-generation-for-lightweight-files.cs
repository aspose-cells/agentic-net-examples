// Title: Batch convert Excel workbooks to lightweight HTML with embedded images using Aspose.Cells in C#
// Common Searches: C# code to convert all .xlsx files in a folder to HTML using Aspose.Cells without generating CSS files | How to export Excel worksheets to single HTML files with embedded Base64 images in Aspose.Cells | Batch processing Excel to HTML in .NET while disabling external CSS with Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | Aspose.Cells disable CSS generation | C# batch Excel to HTML conversion | Excel to HTML lightweight export | Aspose.Cells save multiple workbooks as HTML

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates through every .xlsx file in a specified input directory, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 and suppress external CSS, then saves each workbook as an individual HTML file in an output folder, creating the folder if needed and logging successes or errors.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string inputFolder = @"C:\InputExcels";

        // Folder where the generated HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Retrieve all Excel files (adjust the pattern for .xls if needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            try
            {
                // Ensure the source file exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found, skipping: {excelPath}");
                    continue;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options to produce lightweight files
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Embed images as Base64 to keep a single HTML file (optional)
                    ExportImagesAsBase64 = true,

                    // Export all worksheets (set to true if you only need the active sheet)
                    ExportActiveWorksheetOnly = false
                };

                // Build the output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion to HTML completed.");
    }
}
