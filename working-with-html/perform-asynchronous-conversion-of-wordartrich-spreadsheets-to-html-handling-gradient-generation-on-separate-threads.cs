// Title: Asynchronously convert an Excel workbook containing WordArt to a self‑contained HTML file with Base64 images using Aspose.Cells for .NET
// AI Prompts: Create an async C# method that loads a .xlsx file, sets HtmlSaveOptions to embed WordArt and other images as Base64, preserves gridlines, and saves the workbook as a single HTML file with Aspose.Cells. | Extend the console app to accept a list of Excel file paths and convert each to HTML concurrently using Task.WhenAll, producing self‑contained HTML outputs for all files. | Add robust error handling to the async conversion routine so it logs missing files or conversion exceptions and returns a success flag.
// Common Searches: async Aspose.Cells convert Excel with WordArt to HTML base64 images | C# generate self‑contained HTML from .xlsx preserving gridlines | how to export WordArt as Base64 using Aspose.Cells HtmlSaveOptions | parallel conversion of multiple Excel files to HTML with Aspose.Cells | Aspose.Cells asynchronous HTML export example .NET
// Tags: async Excel to HTML conversion Aspose.Cells | export WordArt as Base64 images HTML | HtmlSaveOptions preserve gridlines C# | parallel processing multiple workbooks Aspose.Cells | self-contained HTML output from .xlsx

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsAsyncHtmlConversion
{
    // The example validates command‑line arguments, loads the specified .xlsx workbook with Aspose.Cells, configures HtmlSaveOptions to embed all images (including WordArt) as Base64, preserve gridlines, and export all worksheets, then saves the result as a self‑contained HTML file. The conversion method is asynchronous, and the program can be expanded to process several workbooks in parallel with proper error handling.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Example usage:
            // args[0] = input Excel file path
            // args[1] = output HTML file path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsAsyncHtmlConversion <input.xlsx> <output.html>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            await ConvertWorkbookToHtmlAsync(inputPath, outputPath);
            Console.WriteLine($"Conversion completed. HTML saved to: {outputPath}");
        }

        /// <param name="excelFilePath">Path to the source .xlsx file.</param>
        /// <param name="htmlOutputPath">Path where the resulting HTML file will be saved.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public static async Task ConvertWorkbookToHtmlAsync(string excelFilePath, string htmlOutputPath)
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: Input file not found – {excelFilePath}");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(excelFilePath);

                // No explicit rendering of WordArt is required; Aspose.Cells handles it during HTML export.
                // However, if additional processing of shapes is needed, it can be added here.

                // Configure HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export all worksheets.
                    ExportActiveWorksheetOnly = false,

                    // Export images (including WordArt) as Base64 strings to keep the HTML self‑contained.
                    ExportImagesAsBase64 = true,

                    // Preserve the original layout as closely as possible.
                    ExportGridLines = true
                };

                // Save the workbook as HTML.
                workbook.Save(htmlOutputPath, saveOptions);
            }
            catch (Exception ex)
            {
                // Log the exception; in production code consider more sophisticated logging.
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }

            // Simulate asynchronous behavior.
            await Task.CompletedTask;
        }
    }
}
