using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook with a default worksheet
        Workbook wb = new Workbook();

        // Add sample content to demonstrate the watermark on each page
        for (int i = 0; i < wb.Worksheets.Count; i++)
        {
            wb.Worksheets[i].Cells["A1"].PutValue($"Page {i + 1}");
        }

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 72);
        font.Bold = true;
        font.Italic = true;
        font.Color = Color.Red;

        // Create the watermark with the desired text and font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);
        watermark.HAlignment = TextAlignmentType.Center;   // Horizontal center
        watermark.VAlignment = TextAlignmentType.Center;   // Vertical center
        watermark.Rotation = 45f;                          // Diagonal (cross) orientation
        watermark.Opacity = 0.3f;                          // Semi‑transparent
        watermark.ScaleToPagePercent = 75;                 // Scale relative to page size

        // Apply the watermark via PDF save options
        PdfSaveOptions options = new PdfSaveOptions();
        options.Watermark = watermark;

        // Save the workbook as a PDF with the watermark applied to every page
        wb.Save("ConfidentialStamped.pdf", options);
    }
}

// Author: Aspose.Cells .NET example – confidential watermark implementation.