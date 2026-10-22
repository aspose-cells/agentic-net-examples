// Title: Convert an Excel workbook to HTML with full cell formatting, gridlines, and embedded images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file and saves it as an HTML document while preserving fonts, colors, and gridlines with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to export the entire workbook, keep cell styles, and embed charts and pictures as Base64 strings. | Provide a robust C# example that checks file existence, handles exceptions, and converts Excel to HTML retaining headers, footers, and cell formatting using Aspose.Cells.
// Common Searches: Aspose.Cells .NET export entire workbook to HTML with cell styles and grid lines | C# save Excel as HTML preserving font colors and background using Aspose | How to embed images as Base64 when converting .xlsx to HTML with Aspose.Cells | Export Excel to HTML with gridlines and embedded charts in C# | Aspose.Cells HtmlSaveOptions settings for preserving formatting in HTML output
// Tags: Aspose.Cells HtmlSaveOptions export full workbook | C# convert Excel to HTML with cell formatting | embed images as Base64 Aspose.Cells HTML | preserve gridlines Aspose.Cells HTML export | retain font colors and styles in HTML output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export the whole workbook, keep grid lines, and embed images as Base64, then saves the result as output.html while handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export the entire workbook (not just the active sheet)
                    ExportActiveWorksheetOnly = false,

                    // Preserve grid lines to match Excel view
                    ExportGridLines = true,

                    // Embed images (charts, pictures) as Base64 strings
                    ExportImagesAsBase64 = true
                    // Cell styles and headers/footers are exported by default
                };

                // Save the workbook as an HTML file with the defined options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
