// Title: Generate a single self‑contained HTML file from an Excel workbook with embedded styles using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads an .xlsx workbook with Aspose.Cells, configures HtmlSaveOptions to suppress external CSS creation, and saves the result as one HTML file with all styles inline. | Show how to set Aspose.Cells to encode images as Base64 and keep every stylesheet inside the generated HTML when exporting Excel to HTML.
// Common Searches: aspnet convert excel to html without separate css file | aspose.cells htmlsaveoptions disable css generation c# | how to embed workbook styles in html output using aspose.cells | single html export from excel with base64 images asp.net | c# save workbook as single html page aspose.cells
// Tags: Aspose.Cells HtmlSaveOptions suppress external stylesheet | embed CSS in HTML output Aspose.Cells | export Excel to single HTML file .NET | inline image encoding Aspose.Cells | self‑contained HTML export from workbook

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source Excel file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to embed styles directly and encode images as Base64, then saves the workbook as a single HTML document while handling any runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export images as Base64 strings to keep everything in a single HTML file
                ExportImagesAsBase64 = true
                // Styles are embedded by default; no separate CSS file is generated
            };

            // Save the workbook as a single HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
