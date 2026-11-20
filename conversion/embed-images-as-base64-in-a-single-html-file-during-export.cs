// Title: How to export an Excel workbook to a single HTML file with Base64‑embedded images using Aspose.Cells for .NET
// Common Searches: aspnet c# save workbook as html with images embedded as base64 using aspose.cells | convert xlsx to single html file with inline base64 images in c# | asp.net core export excel to html and embed pictures as base64 strings | c# htmlsaveoptions exportimagesasbase64 example for aspose.cells | how to generate a self‑contained html from excel using aspose.cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to single HTML with embedded images | embed workbook pictures as base64 in HTML | convert .xlsx to self-contained HTML using Aspose.Cells | save workbook as HTML with inline images C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx file with Aspose.Cells, configures HtmlSaveOptions.ExportImagesAsBase64 = true, and saves the workbook as a single HTML document where all pictures are stored as Base64 strings, producing a self‑contained HTML output while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to embed images as Base64
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
                // The OnePagePerSheet option is not available in this version; default behavior will be used.
            };

            // Save the workbook as a single HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
