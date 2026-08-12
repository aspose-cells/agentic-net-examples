// Title: Generate HTML with Base64‑embedded images from Aspose.Cells (C#)
// Description: C# sample that builds a workbook, adds a picture if available, enables HtmlSaveOptions.ExportImagesAsBase64, and saves a single HTML page where every worksheet image is encoded as a Base64 data‑URI.
// Keywords: Aspose.Cells Base64 HTML | ExportImagesAsBase64 C# | embed images in HTML Aspose | C# Aspose.Cells HTML export | self‑contained HTML report
// Common Searches: Aspose.Cells embed images as Base64 when saving to HTML | C# HtmlSaveOptions ExportImagesAsBase64 example | save Excel workbook as HTML with inline images | convert worksheet to HTML with Base64 pictures
// Developer Intent: Create an HTML output that contains worksheet pictures encoded as Base64 strings, eliminating external image files.
// Use Cases: Send a portable HTML report via email without attaching separate image files. | Publish Excel‑based documentation on a web server where images must stay inline. | Generate web‑ready pages for environments that block external resource loading.
// AI Prompts: Show how to embed pictures from all worksheets as Base64 in the generated HTML. | Provide code that forces UTF‑8 encoding while using ExportImagesAsBase64 in Aspose.Cells. | Explain strategies for handling very large images to keep the HTML size manageable when using Base64 embedding.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlBase64Demo
{
    // C# sample that builds a workbook, adds a picture if available, enables HtmlSaveOptions.ExportImagesAsBase64, and saves a single HTML page where every worksheet image is encoded as a Base64 data‑URI.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (empty)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "example.jpg";

                // Add image if the file exists
                if (File.Exists(imagePath))
                {
                    worksheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Set HTML save options to embed images as Base64
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file
                string outputPath = "output.html";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"HTML file saved to '{outputPath}' with images embedded as Base64.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
