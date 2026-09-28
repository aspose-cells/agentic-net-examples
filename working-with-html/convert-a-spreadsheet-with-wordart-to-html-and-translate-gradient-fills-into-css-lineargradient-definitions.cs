// Title: Convert an Excel workbook containing WordArt to a single‑file HTML document with CSS linear‑gradient for gradient fills using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to embed all images as Base64, and saves the workbook as a single HTML file while ensuring WordArt gradient fills are rendered as CSS linear‑gradient. | Modify the conversion to output separate image files instead of Base64 encoding, and still map WordArt gradient fills to CSS linear‑gradient in the resulting HTML. | Add comprehensive error handling that validates the input path, catches conversion exceptions, and writes detailed logs to a file.
// Common Searches: Aspose.Cells .NET export Excel with WordArt to HTML single‑page | how to preserve gradient fills from Excel WordArt in HTML using Aspose | Aspose.Cells generate HTML with inline data URIs from Excel | convert Excel shape shading to CSS gradients with Aspose.Cells | save Excel as HTML without external image files using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# convert Excel to HTML with WordArt | gradient fill to CSS linear-gradient Aspose.Cells | single-file HTML output from workbook | preserve WordArt formatting in HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook, sets HtmlSaveOptions.ExportImagesAsBase64 to true so that WordArt and other images are embedded, and saves the file as a single HTML page where Excel gradient fills are automatically emitted as CSS linear‑gradient rules.
class SpreadsheetToHtmlConverter
{
    static void Main()
    {
        // Path to the source Excel file that contains WordArt objects.
        string inputFile = "input.xlsx";

        // Path where the generated HTML file will be saved.
        string outputFile = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file not found at '{inputFile}'.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Embed images (including WordArt) as Base64 strings to keep a single HTML file.
                ExportImagesAsBase64 = true
                // WordArt and gradient fills are handled automatically by the library.
            };

            // Save the workbook as an HTML file using the configured options.
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"HTML file successfully generated at '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
