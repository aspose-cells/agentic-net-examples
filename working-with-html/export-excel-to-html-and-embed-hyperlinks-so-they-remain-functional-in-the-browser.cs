// Title: Convert an Excel workbook to HTML with active hyperlinks and embedded images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, configures HtmlSaveOptions to keep hyperlinks clickable, and saves the workbook as an HTML file with images embedded as base64. | Show how to use Aspose.Cells HtmlSaveOptions to export all worksheets to a single HTML page while preserving hyperlink functionality.
// Common Searches: asp.net convert xlsx to html preserving hyperlink clicks | c# Aspose.Cells export workbook to html with base64 images | how to keep Excel hyperlinks working after saving as html using Aspose.Cells | save multiple worksheets to one html file Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions export hyperlinks | C# export Excel to HTML with base64 images | preserve Excel hyperlinks in HTML output | convert multiple worksheets to single HTML page | Aspose.Cells workbook to HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input.xlsx file, loads it with Aspose.Cells, sets HtmlSaveOptions to export all worksheets, embed images as base64, and retain functional hyperlinks, then saves the result as output.html.
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
                // Export all worksheets (set to true to export only the active sheet)
                ExportActiveWorksheetOnly = false,

                // Embed images directly into the HTML as base64 strings
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an HTML file with functional hyperlinks
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
