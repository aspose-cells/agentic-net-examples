using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class AddDiagonalWatermark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample content (optional)
        sheet.Cells["A1"].PutValue("Sample content");

        // Create a rendering font: Times New Roman, 24‑point size
        RenderingFont font = new RenderingFont("Times New Roman", 24);

        // Create a text watermark with the specified font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            Rotation = 45f,                 // Diagonal orientation
            Opacity = 0.3f,                 // 30% opacity
            IsBackground = true,           // Place behind page contents
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            ScaleToPagePercent = 100       // Full page width
        };

        // Assign the watermark to PDF save options
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF to visualize the watermark
        workbook.Save("DiagonalWatermark.pdf", options);
    }
}