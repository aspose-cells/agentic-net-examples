// Title: Batch convert multiple Excel workbooks containing WordArt to HTML with external image files using Aspose.Cells for .NET
// Common Searches: how to batch convert Excel workbooks with WordArt to HTML using Aspose.Cells in C# | aspnet convert all .xlsx files in a folder to HTML with external images | Aspose.Cells export WordArt graphics as separate image files when saving to HTML | C# loop to process multiple Excel files and save each as HTML using HtmlSaveOptions | disable base64 image embedding in Aspose.Cells HTML export
// Tags: batch excel to html conversion Aspose.Cells | export wordart as external images html | HtmlSaveOptions ExportImagesAsBase64 false | process multiple xlsx files c# | save workbook as html external resources

using System;
using System.IO;
using Aspose.Cells;

// The example iterates through every .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to write images (including WordArt) as separate files instead of embedding them as base64, and saves each workbook as an individual HTML file in an output folder, with error handling for missing files and I/O issues.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcels";

        // Folder where the HTML files will be written
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: '{inputFolder}'.");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        try
        {
            // Process each .xlsx file in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(excelPath))
                    continue;

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        // Export images as separate files (not embedded as base64)
                        ExportImagesAsBase64 = false
                    };

                    // Determine the output HTML file path
                    string htmlPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".html");

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);
                }
                catch (Exception ex)
                {
                    // Log any errors for the current file and continue processing remaining files
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors (e.g., access issues) and exit
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
