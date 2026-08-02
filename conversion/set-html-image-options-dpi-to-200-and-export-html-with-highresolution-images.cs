// Title: C# – Export Excel to HTML with 200 DPI Images Using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, optionally add a PNG picture, configure HtmlSaveOptions and ImageOrPrintOptions to export HTML with 200 dpi images saved as separate files (Base64 disabled), and save the result as a high‑resolution HTML document.
// Keywords: Aspose.Cells | HtmlSaveOptions | ImageOrPrintOptions | 200 DPI | high resolution HTML export | C# | .NET | export Excel to HTML | separate image files | disable Base64
// Common Searches: Aspose.Cells set image DPI for HTML export | C# export Excel as HTML with high‑resolution images | HtmlSaveOptions 200 dpi images | How to save HTML images as files instead of Base64 in Aspose.Cells | Increase image quality when converting Excel to HTML
// Developer Intent: Configure Aspose.Cells to generate HTML where embedded images are rendered at 200 dpi and stored as external image files.
// Use Cases: Producing printable web reports that require 200 dpi PNGs for charts and diagrams. | Creating e‑learning or documentation pages where image clarity is critical after Excel‑to‑HTML conversion. | Exporting interactive dashboards to HTML while preserving zoom‑ready image quality.
// AI Prompts: Show how to export images as JPEG with 300 dpi using Aspose.Cells in C#. | Provide code that sets different horizontal and vertical DPI values for HTML export. | Explain how to embed the generated HTML and its external image files into a responsive web page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, optionally add a PNG picture, configure HtmlSaveOptions and ImageOrPrintOptions to export HTML with 200 dpi images saved as separate files (Base64 disabled), and save the result as a high‑resolution HTML document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample content
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("High‑DPI HTML Export");

            // Add an image to the worksheet if the file exists
            string imagePath = "example.png";
            if (File.Exists(imagePath))
            {
                // topRow=2, leftColumn=0 corresponds to cell A3
                sheet.Pictures.Add(2, 0, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set high DPI for exported images
            ImageOrPrintOptions imgOptions = htmlOptions.ImageOptions;
            imgOptions.HorizontalResolution = 200;
            imgOptions.VerticalResolution = 200;

            // Export images as separate files (not Base64) to preserve high resolution
            htmlOptions.ExportImagesAsBase64 = false;

            // Save the workbook as HTML
            string outputPath = "HighDPI_Output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
