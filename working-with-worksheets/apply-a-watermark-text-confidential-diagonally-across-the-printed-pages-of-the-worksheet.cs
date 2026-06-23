using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WatermarkDemo
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to illustrate the watermark effect
        sheet.Cells["A1"].PutValue("Sample data");
        sheet.Cells["A2"].PutValue("Printed with Confidential watermark");

        // Define a rendering font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create a text watermark with the desired properties
        RenderingWatermark watermark = new RenderingWatermark("Confidential", font)
        {
            Rotation = 45f,                     // Diagonal orientation
            Opacity = 0.3f,                     // Semi‑transparent
            IsBackground = true,                // Place behind page content
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            ScaleToPagePercent = 100            // Scale to fit the page
        };

        // Set up PDF save options and assign the watermark
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF; the watermark will appear on each printed page
        workbook.Save("ConfidentialWatermarked.pdf", saveOptions);
    }
}