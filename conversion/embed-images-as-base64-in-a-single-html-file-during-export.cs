// Title: Export an Aspose.Cells workbook to a single self-contained HTML file with base64‑encoded images using C#
// Common Searches: C# Aspose.Cells export Excel to HTML with images embedded as base64 | How to generate a single HTML file from a workbook using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example in .NET | Save Excel worksheet as HTML with embedded PNG picture in C# | Create self-contained HTML from Aspose.Cells workbook without external image files
// Tags: HtmlSaveOptions ExportImagesAsBase64 C# | Aspose.Cells single HTML export with embedded images | self-contained HTML generation from Excel workbook | embed picture in HTML using Aspose.Cells | C# convert workbook to base64 image HTML

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, adds text, numbers, and optionally a PNG picture, configures HtmlSaveOptions.ExportImagesAsBase64 to embed images as base64 strings, and saves the result as a single HTML file that contains all images inline.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["B2"].PutValue(12345);

            // Optional: insert an image if the file exists
            string imagePath = "logo.png";
            if (File.Exists(imagePath))
            {
                // Add picture at row 3, column 1 (zero‑based indices)
                sheet.Pictures.Add(3, 1, imagePath);
            }

            // Configure HTML save options to embed images as base64
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.ExportImagesAsBase64 = true;   // Embed images directly in the HTML

            // Export the workbook to a single HTML file with embedded images
            string outputPath = "ExportedDocument.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
