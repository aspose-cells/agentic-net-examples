// Title: Export an Excel worksheet to HTML with Base64‑embedded images using Aspose.Cells for .NET (prepare for mobile‑friendly styling)
// AI Prompts: Generate C# code that loads a .xlsx file, verifies its existence, and saves the active worksheet as an HTML file with images embedded as Base64 strings using Aspose.Cells HtmlSaveOptions. | Demonstrate how to inject a viewport meta tag and minimal responsive CSS into the HTML produced by Aspose.Cells while preserving the Base64 image embedding. | Provide robust error‑handling patterns for the Excel‑to‑HTML conversion, covering file‑not‑found and generic exceptions.
// Common Searches: how to export only the active sheet to html with aspose.cells c# | aspose.cells embed images as base64 in html output | generate mobile friendly html from excel using asp.net and aspose.cells | add responsive meta viewport to html saved from workbook aspose.cells | c# save workbook as html with base64 images and custom css
// Tags: Aspose.Cells HtmlSaveOptions Base64 images | export active worksheet to HTML .NET | mobile‑optimized HTML from Excel workbook | responsive CSS for Aspose.Cells HTML output | C# Excel to HTML conversion with embedded images

using System;
using System.IO;
using Aspose.Cells;

// The program checks for input.xlsx, loads it with Aspose.Cells, and saves the active worksheet as an HTML file that embeds worksheet images as Base64 strings using HtmlSaveOptions, ready for additional responsive styling.
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
            // Load the source Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export only the active worksheet
                ExportActiveWorksheetOnly = true,
                // Embed images as Base64 strings for easier mobile rendering
                ExportImagesAsBase64 = true
                // Note: Responsive layout is not available in this version of Aspose.Cells.
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"HTML file successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
