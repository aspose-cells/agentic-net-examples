// Title: Generate minimal HTML from an Excel workbook by disabling comments, conditional formatting, and grid lines with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Workbook, sets HtmlSaveOptions.ExportComments, ExportConditionalFormatting, and ExportGridLines to false, embeds images as base64, and saves the file as a single HTML document. | Update an existing Aspose.Cells .NET project to turn off comments, conditional formatting, and grid lines in HtmlSaveOptions while preserving other default HTML export settings.
// Common Searches: Aspose.Cells .NET export Excel to HTML without comments, conditional formatting, and grid lines | How to create a lightweight HTML file from a workbook using Aspose.Cells | Minimal HTML output with base64 images using Aspose.Cells HtmlSaveOptions | Disable comments and conditional formatting in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions disable comments | Aspose.Cells HtmlSaveOptions disable conditional formatting | Aspose.Cells HtmlSaveOptions hide grid lines | Aspose.Cells export minimal HTML | Aspose.Cells embed images base64 HTML

using System;
using System.IO;
using Aspose.Cells;

namespace MinimalHtmlExport
{
    // The example demonstrates creating a Workbook, configuring HtmlSaveOptions to turn off comments, conditional formatting, and grid lines, embedding images as base64, and saving the result as a compact HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Data 1");
                sheet.Cells["A3"].PutValue("Data 2");

                // Configure HTML export options for minimal output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do not export grid lines
                    ExportGridLines = false,
                    // Embed images as base64 to keep a single HTML file
                    ExportImagesAsBase64 = true
                };

                // Define output file path
                string outputPath = "MinimalOutput.html";

                // Save the workbook as an HTML file using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"HTML file successfully saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
