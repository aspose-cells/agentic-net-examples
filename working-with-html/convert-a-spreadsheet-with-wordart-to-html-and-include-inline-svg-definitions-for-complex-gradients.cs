// Title: Export Excel workbook containing WordArt to HTML with embedded Base64 images and inline SVG gradients using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, checks its existence, and saves it as an HTML file using Aspose.Cells while embedding all pictures as Base64 data URIs. | Show how to configure Aspose.Cells HtmlSaveOptions to produce inline SVG elements for WordArt and gradient fills when converting a workbook to HTML. | Create a complete console application that validates the input Excel path, applies the HTML export settings for Base64 images and SVG gradients, and writes the resulting HTML to a specified output file.
// Common Searches: how to export Excel with WordArt to HTML using Aspose.Cells C# | Aspose.Cells HTML export embed images as base64 and preserve gradients | convert Excel gradient fills to inline SVG in .NET | C# save workbook as HTML with WordArt and SVG definitions Aspose
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# Aspose.Cells workbook to HTML conversion | inline SVG generation for Excel WordArt | preserve gradient fills in HTML export Aspose.Cells | embed Excel graphics as data URI in HTML

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The program checks for the existence of an input .xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to embed all images as Base64 data URIs and to render WordArt and complex gradient fills as inline SVG, then saves the workbook as an HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Embed images directly in the HTML as Base64 strings
                ExportImagesAsBase64 = true
            };

            // Save the workbook to HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
