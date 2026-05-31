using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to ensure multiple pages if needed
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.Cells["A1"].PutValue($"Worksheet {i + 1}");
        }

        // Define a large, bold font for the stamp
        RenderingFont font = new RenderingFont("Arial", 72)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark with the desired appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Center the watermark on each page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,

            // Place it over the page content (foreground)
            IsBackground = false,

            // Optional rotation for visual effect
            Rotation = 45,

            // Make it semi‑transparent
            Opacity = 0.3f,

            // Scale relative to the page size
            ScaleToPagePercent = 75
        };

        // Configure PDF save options to use the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF; the watermark will appear on every page
        workbook.Save("ConfidentialStamped.pdf", pdfOptions);
    }
}