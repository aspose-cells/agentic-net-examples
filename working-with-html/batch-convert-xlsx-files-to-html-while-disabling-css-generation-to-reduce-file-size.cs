// Title: Batch convert XLSX workbooks to HTML with Aspose.Cells in C# while disabling CSS to shrink output size
// AI Prompts: Create a C# console program that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file in a target directory. | Configure the HTML save options to turn off embedded CSS (ExportEmbeddedCss = false) and optionally embedded images to minimize the generated HTML file size. | Add try‑catch logging that records the path of any workbook that fails to convert and continues processing the remaining files.
// Common Searches: c# batch convert multiple xlsx files to html using aspose.cells without css | how to turn off css generation when exporting excel to html with aspose.cells | reduce size of html files produced by aspose.cells excel conversion | process all spreadsheets in a folder and save as html in .net core | asp.net console app convert excel to html without embedded styles
// Tags: aspose.cells exportembeddedcss false | c# excel to html conversion without embedded css | reduce html file size aspose.cells | process multiple excel workbooks programmatically c# | xlsx to html conversion aspnet

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsxToHtml
{
    // // This C# console application iterates through all .xlsx files in a given input folder, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output folder. HtmlSaveOptions are set with ExportEmbeddedCss = false (and optionally ExportEmbeddedImages = false) to prevent CSS generation, resulting in smaller HTML files. The program creates the output directory if missing and logs successful conversions and any errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing XLSX files
            string inputFolder = @"C:\InputXlsx";
            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtml";

            // Ensure input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            // Configure HTML save options (defaults export all sheets with inline styles)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Verify the source file exists
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook from the XLSX file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Determine output HTML file name (same base name as XLSX)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
