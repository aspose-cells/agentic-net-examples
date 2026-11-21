// Title: Convert an Excel workbook to HTML with full cell formatting using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook to HTML preserving fonts and colors | How to keep Excel cell styles when saving as HTML with Aspose.Cells | Save Excel as HTML with gridlines and embedded images using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions to retain cell formatting in HTML output | Convert .xlsx to .html while preserving cell formatting in C#
// Tags: HtmlSaveOptions PreserveFormatting | Aspose.Cells HTML export with base64 images | C# Excel to HTML conversion retaining fonts | gridlines inclusion in HTML output Aspose.Cells | full worksheet export to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing .xlsx file with Aspose.Cells, configures HtmlSaveOptions to export all worksheets, embed images as Base64, keep gridlines and cell styles, and saves the result as an .html file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to preserve cell formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false,   // Export all worksheets
                ExportImagesAsBase64 = true,         // Embed images directly
                ExportGridLines = true               // Keep grid lines visible
                // Note: Properties such as ExportColumnHeaders, ExportRowHeaders,
                // FontEmbeddingMode, and PreserveFormatting are not available in
                // the current Aspose.Cells version and have been omitted.
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
