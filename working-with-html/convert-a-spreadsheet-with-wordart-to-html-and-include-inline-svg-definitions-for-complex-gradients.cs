// Title: Convert an Excel workbook with WordArt to HTML and embed WordArt as inline SVG using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel WordArt to HTML with inline SVG | How to keep WordArt gradients when saving Excel as HTML in .NET | Aspose.Cells HtmlSaveOptions for embedding WordArt as SVG | Export entire workbook with WordArt to HTML using Base64 images in C# | Convert Excel file containing WordArt to HTML with SVG definitions
// Tags: Aspose.Cells HtmlSaveOptions export WordArt | inline SVG WordArt conversion .NET | Excel to HTML Base64 images Aspose | preserve WordArt gradients Aspose.Cells | C# export workbook as HTML including WordArt

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For HtmlSaveOptions
using System.Drawing.Imaging; // For ImageFormat (if needed)

// The example loads an Excel workbook that contains WordArt objects, configures HtmlSaveOptions to embed images as Base64 and to export the full workbook, then saves it as HTML. Supported WordArt shapes are rendered as inline SVG definitions, preserving complex gradients.
class SpreadsheetToHtmlWithWordArt
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithWordArt.xlsx";
            const string outputPath = "OutputWithWordArt.html";

            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook that contains WordArt objects.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Embed images (including any raster parts of WordArt) as Base64 strings.
                ExportImagesAsBase64 = true,

                // Export the entire workbook (set to true to export only the active sheet).
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as an HTML file. The resulting HTML will contain
            // inline <svg> definitions for WordArt where supported.
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
