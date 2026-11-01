// Title: Export an Excel workbook to a single self‑contained HTML file with Base64‑encoded images using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx as a single HTML page with embedded images using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example for .NET | C# generate HTML from Excel without external CSS or image files using Aspose.Cells | convert Excel workbook to HTML with images inlined as Base64 strings in C# | Aspose.Cells export Excel to HTML single file option
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to single HTML file | inline Excel images as Base64 in HTML | Aspose.Cells workbook to HTML conversion | self‑contained HTML output from Excel

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xlsx exists, loads it with Aspose.Cells, sets HtmlSaveOptions to embed images as Base64 strings, and saves the workbook as a single HTML file (output.html).
class ExcelToHtmlWithConditionalFormatting
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string htmlOutputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export images as Base64 strings
                ExportImagesAsBase64 = true
                // By default, Aspose.Cells creates a separate CSS file when ExportToSingleFile is false
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlOutputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to \"{htmlOutputPath}\"");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
