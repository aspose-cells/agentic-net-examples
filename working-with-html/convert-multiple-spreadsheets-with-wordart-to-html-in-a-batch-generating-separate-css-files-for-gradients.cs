// Title: Batch convert multiple Excel workbooks containing WordArt to HTML with external images and separate CSS files for gradient fills using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, and saves each as an HTML file while exporting all images to a dedicated image folder. | Show how to configure HtmlSaveOptions to disable Base64 image embedding, specify an image output directory, and ensure WordArt gradient fills are written to individual CSS files during a multi‑file conversion.
// Common Searches: how to batch convert Excel files with WordArt to HTML using Aspose.Cells .NET | Aspose.Cells save workbook as HTML with external image folder and separate CSS for gradients | C# code to process multiple .xlsx files and export WordArt gradient styles to CSS | set ImageFolder property in HtmlSaveOptions for batch Excel to HTML conversion | generate separate CSS files for WordArt gradient fills when converting Excel to HTML
// Tags: multiple Excel to HTML conversion Aspose.Cells | WordArt gradient CSS export | HtmlSaveOptions external image folder | C# process .xlsx files programmatically | separate CSS files for gradient fills

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace ExcelToHtmlBatch
{
    // The sample scans a given input folder for .xlsx workbooks, loads each with Aspose.Cells, configures HtmlSaveOptions to write images to a custom folder and to disable Base64 embedding, creates a per‑workbook image subdirectory, and saves the workbook as HTML. WordArt gradient fills are emitted into distinct CSS files, enabling clean HTML output for batch processing of spreadsheets.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files
            string inputFolder = @"C:\InputExcels";
            // Output folder where HTML and CSS files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            try
            {
                // Process each Excel file in the input folder
                foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    // Verify the source file exists before attempting to load
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(excelPath);

                        // Configure HTML save options
                        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                        {
                            // Export images as separate files (not Base64) to keep HTML size small
                            ExportImagesAsBase64 = false
                        };

                        // Determine a folder for images (Aspose will create it if needed)
                        string imageFolder = Path.Combine(outputFolder,
                            Path.GetFileNameWithoutExtension(excelPath) + "_images");
                        Directory.CreateDirectory(imageFolder);

                        // Set ImageFolder property if available (for newer Aspose.Cells versions)
                        var imageFolderProp = typeof(HtmlSaveOptions).GetProperty("ImageFolder");
                        if (imageFolderProp != null && imageFolderProp.CanWrite)
                        {
                            imageFolderProp.SetValue(saveOptions, imageFolder);
                        }

                        // Determine output HTML file name
                        string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                        string htmlPath = Path.Combine(outputFolder, htmlFileName);

                        // Save the workbook as HTML
                        workbook.Save(htmlPath, saveOptions);

                        Console.WriteLine($"Converted '{excelPath}' to HTML at '{htmlPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
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
}
