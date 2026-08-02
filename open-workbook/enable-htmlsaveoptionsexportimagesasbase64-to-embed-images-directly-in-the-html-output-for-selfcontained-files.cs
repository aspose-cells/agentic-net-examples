// Title: Aspose.Cells C# – Export Worksheet Images as Base64 in a Self‑Contained HTML File
// Description: Demonstrates how to insert a picture into a worksheet, enable HtmlSaveOptions.ExportImagesAsBase64, and save the workbook as a single HTML document with all images embedded as Base64 data URIs, eliminating external image files.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportImagesAsBase64 | embed images base64 | self‑contained HTML | Excel to HTML conversion | picture insertion | code example | GitHub
// Common Searches: Aspose.Cells ExportImagesAsBase64 example | C# save Excel as HTML with embedded images | HtmlSaveOptions ExportImagesAsBase64 .NET tutorial | self‑contained HTML from workbook Aspose | embed Excel picture as Base64 in HTML
// Developer Intent: Create an HTML file from a workbook where every worksheet image is encoded as a Base64 string.
// Use Cases: Generate portable HTML reports that include logos or charts without external files. | Embed branded Excel templates directly into email bodies or web pages. | Provide offline‑viewable HTML versions of spreadsheets where images stay inside the document.
// AI Prompts: Write C# code that adds a picture to a worksheet and saves the workbook as HTML with ExportImagesAsBase64 enabled using Aspose.Cells. | Explain how HtmlSaveOptions.ExportImagesAsBase64 affects the HTML output and file size. | Give troubleshooting steps when Base64‑encoded images are missing from the generated HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to insert a picture into a worksheet, enable HtmlSaveOptions.ExportImagesAsBase64, and save the workbook as a single HTML document with all images embedded as Base64 data URIs, eliminating external image files.
    public class ExportImagesAsBase64Demo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "example.jpg";

                // Add image if file exists
                if (File.Exists(imagePath))
                {
                    worksheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping image insertion.");
                }

                // Create HTML save options with Base64 image embedding
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as a self‑contained HTML file
                string outputPath = "output.html";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"HTML file '{outputPath}' with embedded Base64 images saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportImagesAsBase64Demo.Run();
        }
    }
}
