// Title: Batch convert multiple .xlsx workbooks to HTML with embedded base64 images using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file using HtmlSaveOptions with ExportImagesAsBase64 enabled. | Add code to verify that the input folder exists, create the output folder if it does not, and log a success or error message for every workbook processed. | Configure HtmlSaveOptions so that the generated HTML file name matches the source workbook name (without extension) and all worksheet images are embedded as base64 strings.
// Common Searches: how to batch convert xlsx files to html with aspose.cells in c# | c# aspose.cells export images as base64 in html output for multiple workbooks | asp.net core program to convert a folder of excel files to html with embedded images | example of HtmlSaveOptions ExportImagesAsBase64 for batch processing Excel files | c# console app to generate html files from excel workbooks using Aspose.Cells
// Tags: batch excel to html conversion Aspose.Cells | export images as base64 HtmlSaveOptions | c# console workbook folder processing | aspose.cells htmlsaveoptions configuration | generate html filenames from workbook names

using System;
using System.IO;
using Aspose.Cells;

// A C# console program iterates over all .xlsx files in an input directory, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output directory. HtmlSaveOptions are configured with ExportImagesAsBase64=true, and the code ensures the output folder exists and logs conversion results.
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Input folder containing Excel files
        string inputFolder = @"C:\InputExcelFiles";
        // Output folder for generated HTML files
        string outputFolder = @"C:\OutputHtmlFiles";

        // Verify input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (adjust the pattern as needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string excelPath in excelFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportImagesAsBase64 = true
                };

                // Determine output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
