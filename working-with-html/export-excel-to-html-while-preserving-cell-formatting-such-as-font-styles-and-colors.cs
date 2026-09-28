// Title: Convert an Excel workbook to HTML with full cell formatting (fonts, colors, and gridlines) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file (or creates a new workbook when the file is absent) and saves it as an HTML document with Aspose.Cells, preserving font styles, text colors, and gridlines. | Show how to set up Aspose.Cells HtmlSaveOptions to embed worksheet images as Base64 strings and to output all worksheets into one HTML file. | Add comprehensive try‑catch error handling that writes any conversion exceptions to the console.
// Common Searches: Aspose.Cells C# export Excel workbook to HTML with original text colors | retain cell borders when converting .xlsx to HTML using Aspose.Cells | save Excel as HTML and embed pictures directly in the page with Aspose.Cells | merge all worksheets into a single HTML document with Aspose.Cells .NET | example code for converting Excel to HTML while keeping cell formatting in C#
// Tags: Aspose.Cells HtmlSaveOptions styling retention | preserve text styling in HTML output | cell border rendering in HTML export | inline image encoding for HTML export Aspose.Cells | combine multiple worksheets into one HTML page

using System;
using System.IO;
using Aspose.Cells;

// The C# program checks for 'input.xlsx', loads it or creates a new workbook with sample data, configures HtmlSaveOptions to export all worksheets, retain gridlines, embed images as Base64, and keep font styles and colors, then saves the result as 'output.html' and logs success or any errors to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Load the workbook if the file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add minimal content to avoid an empty workbook.
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false, // Export all worksheets.
                ExportGridLines = true,            // Keep grid lines visible.
                ExportImagesAsBase64 = true        // Embed images directly in HTML.
                // Note: ExportColumnHeaders, ExportRowHeaders, and FontEmbeddingMode are not
                // available in the current Aspose.Cells version and have been omitted.
            };

            // Save the workbook as an HTML file.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
