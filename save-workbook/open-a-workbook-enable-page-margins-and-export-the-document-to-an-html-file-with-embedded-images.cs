// Title: Export an Excel workbook to HTML with embedded Base64 images using Aspose.Cells for .NET
// Common Searches: how to export an Excel workbook to HTML with embedded Base64 images using Aspose.Cells in C# | Aspose.Cells C# example for saving XLSX as HTML with images encoded as Base64 | convert XLSX file to HTML while preserving page layout with Aspose.Cells .NET | C# code to generate HTML from Excel and embed pictures as Base64 strings using Aspose.Cells | Aspose.Cells HTMLSaveOptions ExportImagesAsBase64 usage in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export workbook to HTML with embedded images | Aspose.Cells save workbook as HTML base64 images | Aspose.Cells HTML export page margin support | convert XLSX to HTML using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX workbook with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 strings via ExportImagesAsBase64, notes that page‑margin export is not supported in the current API, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Embed images directly into the HTML as Base64 strings
                ExportImagesAsBase64 = true
                // Note: ExportPageMargins is not available in the current HtmlSaveOptions API
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
