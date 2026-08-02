using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content");

        // Create a rendering font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36);

        // Initialize the text watermark with the desired text and font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);

        // Position the watermark at coordinates (50, 400) on each page
        watermark.HAlignment = TextAlignmentType.Left;   // Align to left edge
        watermark.VAlignment = TextAlignmentType.Top;    // Align to top edge
        watermark.OffsetX = 50;                          // Horizontal offset
        watermark.OffsetY = 400;                         // Vertical offset

        // Optional appearance settings
        watermark.Opacity = 0.3f;        // Semi‑transparent
        watermark.IsBackground = true;  // Place behind page contents

        // Apply the watermark via PDF save options
        PdfSaveOptions options = new PdfSaveOptions();
        options.Watermark = watermark;

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("output.pdf", options);
    }
}

// Author: Aspose.Cells .NET example – adds a positioned text watermark to each PDF page.