// Title: Export an Excel workbook to a single HTML file with embedded Base64 images using Aspose.Cells for .NET (conditional formatting not supported)
// Common Searches: Aspose.Cells C# export workbook to HTML with images embedded as Base64 | How to generate a single HTML file from .xlsx using Aspose.Cells .NET | Save Excel as HTML with embedded images using Aspose.Cells SaveOptions | Export Excel to HTML without external image files Aspose.Cells C# | Aspose.Cells HTML export conditional formatting not available
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to single HTML file | Aspose.Cells HTML export embedded images | Excel to HTML conversion without external image files | conditional formatting export limitation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to embed images as Base64 (producing a single HTML file), notes that ExportConditionalFormatting and CssClassPrefix are unavailable in the current version, saves the workbook as HTML, and catches any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export images as Base64 to keep a single HTML file
                ExportImagesAsBase64 = true
                // Note: ExportConditionalFormatting and CssClassPrefix are not available
                // in the current Aspose.Cells version; they are omitted.
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
