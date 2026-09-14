// Title: Batch convert a folder of .xlsx workbooks to HTML with external image files using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions with ExportImagesAsBase64 set to false. | Create a .NET script that ensures an output folder exists, processes every Excel file in an input folder, configures HtmlSaveOptions to export images as separate files, and logs conversion results or errors. | Generate example code showing how to use Aspose.Cells HtmlSaveOptions to batch convert multiple Excel workbooks to HTML while keeping images external for CSS‑based deduplication.
// Common Searches: how to batch convert xlsx files to html with aspose.cells c# | c# console app to export excel workbooks as html without base64 images | asp.net script for converting a folder of excel files to html using aspose.cells | save excel as html with external image files using Aspose.Cells HtmlSaveOptions
// Tags: batch excel to html conversion Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 false | external image files from excel html export | c# console application Aspose.Cells | css custom properties image deduplication Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates over all .xlsx files in a given input directory, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to keep images as separate files (ExportImagesAsBase64 = false), and saves each workbook as an HTML file in an output folder, handling missing files and logging success or errors.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all Excel files (adjust the pattern if you need .xls files as well)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string excelPath in excelFiles)
            {
                // Verify the source file exists before attempting to load
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the Excel workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Set HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Keep images as separate files (set to true if you prefer Base64)
                        ExportImagesAsBase64 = false
                    };

                    // Build the output HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, saveOptions);
                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
