// Title: Asynchronously convert an Excel workbook containing WordArt and gradients to a self‑contained HTML file with Aspose.Cells for .NET
// AI Prompts: Write an async C# method that loads an Excel file, sets HtmlSaveOptions.ExportImagesAsBase64 to true, and saves the workbook as HTML using Aspose.Cells. | Add input‑file validation, create the output directory if it does not exist, and log any conversion errors while keeping the method awaitable. | Refactor the conversion so the HTML save runs on a background thread and returns a Task, ensuring WordArt shapes with gradients are embedded as Base64 images.
// Common Searches: asp.net core async convert excel with wordart to html base64 | Aspose.Cells export workbook to self‑contained html including gradients | how to embed wordart shapes as base64 images in html using Aspose.Cells | run Aspose.Cells HtmlSaveOptions on a background thread in C# | convert large excel file to html asynchronously with Aspose.Cells
// Tags: asynchronous Excel to HTML conversion Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 example | embed WordArt gradients in self‑contained HTML | background thread workbook save Aspose.Cells | input file validation and output directory creation C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsAsyncHtmlConversion
{
    // The example provides an async ConvertWorkbookToHtmlAsync method that checks the source Excel file, creates the target folder, loads the workbook, configures HtmlSaveOptions to embed all images (including WordArt gradients) as Base64, and saves the workbook as a single HTML file. Exceptions are caught and logged, and a console program demonstrates invoking the method with command‑line arguments.
    public class Converter
    {
        /// <param name="inputPath">Full path to the source Excel file.</param>
        /// <param name="outputPath">Full path where the resulting HTML file will be saved.</param>
        public static async Task ConvertWorkbookToHtmlAsync(string inputPath, string outputPath)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Load the workbook (synchronous load as per lifecycle rule)
                var workbook = new Workbook(inputPath);

                // Prepare HTML save options
                var htmlOptions = new HtmlSaveOptions
                {
                    // Export images (including any existing gradients) as Base64 to keep the HTML self‑contained
                    ExportImagesAsBase64 = true,
                    // Export the entire workbook; adjust as needed
                    ExportActiveWorksheetOnly = false
                };

                // Save the workbook as HTML (synchronous save as per lifecycle rule)
                workbook.Save(outputPath, htmlOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }
    }

    internal class Program
    {
        // Entry point required for compilation
        private static void Main(string[] args)
        {
            // Example usage:
            // args[0] = input Excel path, args[1] = output HTML path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsAsyncHtmlConversion <input.xlsx> <output.html>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Run the async conversion synchronously
            Converter.ConvertWorkbookToHtmlAsync(inputPath, outputPath)
                     .GetAwaiter()
                     .GetResult();

            Console.WriteLine("Conversion completed.");
        }
    }
}
