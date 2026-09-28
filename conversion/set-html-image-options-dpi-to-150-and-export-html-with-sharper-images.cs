// Title: Export Excel to HTML with 150 DPI images using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as HTML with image resolution set to 150 DPI using Aspose.Cells. | Show how to configure HtmlSaveOptions.ImageOptions to set HorizontalResolution and VerticalResolution before exporting to HTML.
// Common Searches: how to set image DPI in Aspose.Cells HTML export C# | Aspose.Cells export workbook to HTML with high‑resolution images 150 DPI | C# HtmlSaveOptions ImageOptions HorizontalResolution VerticalResolution example
// Tags: Aspose.Cells HtmlSaveOptions image DPI configuration | C# export workbook to HTML high‑resolution images | set HorizontalResolution VerticalResolution Aspose.Cells | disable Base64 image embedding Aspose.Cells HTML

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // The sample creates a workbook, configures HtmlSaveOptions.ImageOptions to use a 150 DPI horizontal and vertical resolution, disables Base64 image embedding, and saves the workbook as an HTML file with sharper, higher‑resolution images.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample HTML Export with 150 DPI images");
            sheet.Cells["A2"].PutValue("This image will be rendered at higher resolution.");

            // Optionally add an image to demonstrate the effect
            // (Replace the path with a valid image file if needed)
            // sheet.Pictures.Add(2, 0, "example.jpg");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Access the ImageOrPrintOptions through HtmlSaveOptions.ImageOptions
            ImageOrPrintOptions imgOptions = htmlOptions.ImageOptions;

            // Set the desired DPI for both horizontal and vertical resolution
            imgOptions.HorizontalResolution = 150; // 150 DPI horizontally
            imgOptions.VerticalResolution = 150;   // 150 DPI vertically

            // Export images as separate files (optional, improves clarity)
            htmlOptions.ExportImagesAsBase64 = false;

            // Save the workbook as HTML using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with images rendered at 150 DPI.");
        }
    }
}
