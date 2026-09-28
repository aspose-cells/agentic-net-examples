// Title: Convert an Excel workbook to responsive HTML5 with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to generate HTML5, embeds all worksheet images as Base64, and saves a mobile‑optimized HTML file. | Show how to set Aspose.Cells HtmlSaveOptions to export every worksheet and produce responsive HTML suitable for smartphones.
// Common Searches: Aspose.Cells .NET export all worksheets to a single responsive HTML page | How to embed Excel images as Base64 when saving to HTML with Aspose.Cells | Generate mobile‑friendly HTML5 from an .xlsx file using C# | Set HtmlSaveOptions for responsive design in Aspose.Cells export
// Tags: Aspose.Cells HtmlSaveOptions responsive export | export excel to html5 with base64 images | C# generate mobile‑friendly HTML from .xlsx | save all worksheets as single HTML using Aspose.Cells | embed worksheet images as base64 in HTML output

using Aspose.Cells;
using System;
using System.IO;

// // This program checks for input.xlsx, loads it with Aspose.Cells, configures HtmlSaveOptions to produce HTML5, exports all worksheets, embeds images as Base64 strings for mobile friendliness, and saves the result to output.html while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the source Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options for responsive design
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Generate HTML5 markup
                HtmlVersion = HtmlVersion.Html5,

                // Export all worksheets (set to true to export only the active sheet)
                ExportActiveWorksheetOnly = false,

                // Embed images directly as Base64 strings to avoid external image files on mobile devices
                ExportImagesAsBase64 = true
            };

            // Save the workbook as a responsive HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
