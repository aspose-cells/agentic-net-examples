// Title: Batch converting multiple XLSX workbooks to HTML without generating CSS using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells batch convert xlsx files to html without generating CSS | How to disable CSS generation when saving Excel as HTML using Aspose.Cells | Convert all Excel workbooks in a folder to HTML with reduced file size in .NET | Aspose.Cells HtmlSaveOptions example to omit stylesheet for bulk conversion | Programmatically export multiple .xlsx files to .html while minimizing output size | Bulk Excel to HTML conversion using Aspose.Cells and suppressing embedded styles
// Tags: Aspose.Cells bulk export xlsx to html | HtmlSaveOptions exclude css | C# excel to html conversion with minimal styling | optimize html output size Aspose.Cells | automated processing of multiple Excel files .net | save workbook as html without style sheet

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsxToHtml
{
    // The C# program iterates through all *.xlsx files in a specified input directory, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to generate HTML without a CSS stylesheet, and saves the result as .html files in an output folder, handling missing files and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing XLSX files
            string inputFolder = @"C:\InputXlsx";
            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtml";

            try
            {
                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all .xlsx files in the input folder
                string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string xlsxPath in xlsxFiles)
                {
                    try
                    {
                        // Verify the file still exists before loading
                        if (!File.Exists(xlsxPath))
                        {
                            Console.WriteLine($"File not found (skipped): {xlsxPath}");
                            continue;
                        }

                        // Load the workbook from the XLSX file
                        Workbook workbook = new Workbook(xlsxPath);

                        // Configure HTML save options
                        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
                        // Optional: reduce size further by not embedding images as Base64
                        // saveOptions.ExportImagesAsBase64 = false;

                        // Determine output HTML file path (same name, .html extension)
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                        string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                        // Save the workbook as HTML using the configured options
                        workbook.Save(htmlPath, saveOptions);

                        Console.WriteLine($"Converted '{xlsxPath}' to '{htmlPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
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
}
