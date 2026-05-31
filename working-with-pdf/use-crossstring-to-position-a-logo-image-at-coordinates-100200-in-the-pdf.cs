using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            const string logoPath = "logo.png";

            // Verify that the logo file exists before attempting to read it
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                return;
            }

            // Load the logo image into a byte array
            byte[] logoBytes = File.ReadAllBytes(logoPath);

            // Create an image‑based watermark using the logo data
            RenderingWatermark watermark = new RenderingWatermark(logoBytes)
            {
                // Position the watermark at (100, 200) points on the PDF page
                OffsetX = 100f,
                OffsetY = 200f,

                // Make the watermark appear in front of the content
                IsBackground = false
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Create an empty workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Save the workbook as PDF with the positioned logo
            workbook.Save("LogoPositioned.pdf", pdfOptions);

            Console.WriteLine("PDF generated successfully: LogoPositioned.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}