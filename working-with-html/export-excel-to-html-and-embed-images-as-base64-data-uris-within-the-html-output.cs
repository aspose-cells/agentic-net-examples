// Title: Export Excel to HTML with Embedded Base64 Images using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add text and a picture, configure HtmlSaveOptions to embed images as Base64 data URIs, and save the result as a self‑contained HTML file while handling missing image files gracefully.
// Keywords: Aspose.Cells | C# HTML export | Base64 image embedding | ExportImagesAsBase64 | self-contained HTML | Excel to HTML conversion | picture insertion Aspose | HtmlSaveOptions | no external images | Windows .NET
// Common Searches: Aspose.Cells export HTML base64 | C# save workbook as HTML with embedded images | HtmlSaveOptions ExportImagesAsBase64 example | embed Excel picture in HTML using Aspose | generate single file HTML from Excel .NET
// Developer Intent: Produce a single HTML document from an Excel workbook where all inserted pictures are encoded as Base64 data URIs.
// Use Cases: Deliver a portable HTML report that includes Excel data and graphics without separate image files. | Embed a snapshot of an Excel sheet in an email or web page with all visuals inlined. | Store a web‑ready view of a worksheet in a CMS where external assets are not allowed.
// AI Prompts: Show a C# example that saves an Aspose.Cells workbook to HTML with all worksheet images embedded as Base64 data URIs. | Explain how to combine HtmlSaveOptions for Base64 image embedding with custom CSS styling in the generated HTML. | Provide guidance on handling missing picture files when exporting Excel to HTML with embedded images using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // Demonstrates how to create a workbook, add text and a picture, configure HtmlSaveOptions to embed images as Base64 data URIs, and save the result as a self‑contained HTML file while handling missing image files gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data
                sheet.Cells["A1"].PutValue("Sample Text");

                // Path to the image file
                string imagePath = "example.jpg";

                // Insert image if the file exists
                if (File.Exists(imagePath))
                {
                    sheet.Pictures.Add(2, 2, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Set HTML save options to embed images as Base64 strings
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file with embedded images
                workbook.Save("output.html", saveOptions);
                Console.WriteLine("Workbook saved to output.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
