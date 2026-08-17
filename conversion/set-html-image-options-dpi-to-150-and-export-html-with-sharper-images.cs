// Title: Export Excel to HTML with 150 DPI Images Using Aspose.Cells for .NET
// Description: Shows how to set horizontal and vertical image resolution to 150 DPI via HtmlSaveOptions.ImageOptions, disable Base64 embedding, and save a workbook as HTML so that charts and pictures render sharply.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ImageOrPrintOptions | 150 DPI | high‑resolution HTML export | Excel to HTML image quality | disable Base64 images | set image resolution | Aspose.Cells example
// Common Searches: Aspose.Cells set image DPI HTML export | HTML export high resolution images Aspose .NET | How to change image resolution in HtmlSaveOptions | Export Excel to HTML with 150 DPI charts | Aspose.Cells disable Base64 images
// Developer Intent: Configure image DPI to 150 and generate HTML output with separate image files for sharper rendering.
// Use Cases: Web dashboards that display Excel charts on retina or high‑DPI screens. | Automated report generation where image clarity is critical. | Embedding Excel‑derived graphics into web pages without Base64 overhead. | Creating printable HTML versions of workbooks with high‑resolution images.
// AI Prompts: Generate C# code that saves the workbook to a MemoryStream with 150 DPI image settings. | Show how to embed images as Base64 while preserving 150 DPI resolution in Aspose.Cells HTML export. | Explain how to apply the same 150 DPI settings when exporting to PDF or PNG. | Provide a step‑by‑step guide to batch‑process multiple workbooks using these HTML settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Shows how to set horizontal and vertical image resolution to 150 DPI via HtmlSaveOptions.ImageOptions, disable Base64 embedding, and save a workbook as HTML so that charts and pictures render sharply.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample HTML Export with 150 DPI images");
            sheet.Cells["A2"].PutValue("This image will be rendered at higher resolution.");

            // Optionally add an image to demonstrate the DPI effect
            // (Replace the path with a valid image file if needed)
            // sheet.Pictures.Add(2, 0, "example.jpg");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Access the ImageOrPrintOptions through HtmlSaveOptions.ImageOptions
            ImageOrPrintOptions imgOptions = htmlOptions.ImageOptions;

            // Set both horizontal and vertical DPI to 150 for sharper images
            imgOptions.HorizontalResolution = 150;
            imgOptions.VerticalResolution = 150;

            // Export images as separate files (optional, set to true for Base64 embedding)
            htmlOptions.ExportImagesAsBase64 = false;

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("HTML file saved with images rendered at 150 DPI.");
        }
    }
}
