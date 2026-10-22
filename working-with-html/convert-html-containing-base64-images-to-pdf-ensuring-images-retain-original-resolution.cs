// Title: Generate a PDF from HTML containing base64‑encoded images while preserving original resolution using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML string with data‑uri PNG images, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves the workbook as a PDF that keeps the images at their native resolution. | Show how to configure Aspose.Cells PdfSaveOptions to prevent down‑sampling or compression of embedded images during HTML‑to‑PDF conversion. | Add error handling and logic to create the output directory automatically before calling Workbook.Save.
// Common Searches: Aspose.Cells .NET convert HTML string with base64 images to PDF preserving image quality | C# load HTML containing data URI images into Aspose.Cells workbook | How to keep original resolution of base64 PNG images when exporting PDF with Aspose.Cells | PdfSaveOptions settings to avoid image compression in Aspose.Cells HTML to PDF conversion | Create output folder automatically when saving PDF from Aspose.Cells workbook
// Tags: HTML to PDF conversion with base64 images Aspose.Cells | PdfSaveOptions image quality Aspose.Cells .NET | HtmlLoadOptions memory stream C# Aspose.Cells | preserve native image resolution Aspose.Cells PDF export | auto-create output directory Aspose.Cells save

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an HTML string that includes base64‑encoded PNG images into an Aspose.Cells Workbook via a MemoryStream and HtmlLoadOptions, then saves the workbook as a PDF using default PdfSaveOptions which retain the original image resolution, creating the output folder if it does not exist.
class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            // HTML string that contains base64‑encoded images
            string htmlContent = @"<html><body>
                <h1>Sample Document</h1>
                <img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...'/>
                </body></html>";

            // Load the HTML content into a workbook using a memory stream and HtmlLoadOptions
            Workbook workbook;
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(htmlStream, loadOptions);
            }

            // Set PDF save options (default options preserve image quality)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
