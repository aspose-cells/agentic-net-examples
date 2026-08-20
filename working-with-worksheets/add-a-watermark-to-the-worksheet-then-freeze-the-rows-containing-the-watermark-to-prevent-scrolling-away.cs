// Title: Add a Text Watermark and Freeze Rows with Aspose.Cells for .NET
// Description: Demonstrates how to create a large, semi‑transparent diagonal text watermark using RenderingWatermark, assign it to PdfSaveOptions, freeze the rows that contain the watermark with FreezePanes, and save the workbook as a PDF where the watermark stays visible and the frozen rows remain on‑screen while scrolling.
// Keywords: Aspose.Cells watermark .NET | RenderingWatermark PDF | FreezePanes C# | Aspose.Cells add text watermark | freeze rows Aspose.Cells | PDFSaveOptions watermark | Aspose.Cells example | C# Excel to PDF watermark
// Common Searches: Aspose.Cells add diagonal text watermark to PDF | How to freeze rows after applying a watermark in Aspose.Cells | C# code for RenderingWatermark with FreezePanes | Set watermark opacity and rotation in Aspose.Cells | Freeze top rows in Excel workbook using Aspose.Cells
// Developer Intent: Create a PDF with a diagonal text watermark and keep the watermark rows fixed by freezing them.
// Use Cases: Generate confidential PDFs with a light‑gray diagonal watermark while keeping header rows visible for reference. | Apply a background watermark to a workbook and lock the first N rows so they never scroll out of view. | Combine RenderingFont, RenderingWatermark, PdfSaveOptions, and FreezePanes to produce a professionally formatted PDF.
// AI Prompts: Show me C# code to add a semi‑transparent diagonal text watermark to an Aspose.Cells workbook and freeze the first three rows. | Provide an Aspose.Cells example that uses RenderingWatermark with custom font, opacity, and rotation, then applies FreezePanes before saving to PDF. | Explain how to synchronize watermark placement with frozen panes in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a large, semi‑transparent diagonal text watermark using RenderingWatermark, assign it to PdfSaveOptions, freeze the rows that contain the watermark with FreezePanes, and save the workbook as a PDF where the watermark stays visible and the frozen rows remain on‑screen while scrolling.
class WatermarkFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional)
        sheet.Cells["A1"].PutValue("Row 1");
        sheet.Cells["A2"].PutValue("Row 2");
        sheet.Cells["A3"].PutValue("Row 3");
        sheet.Cells["A4"].PutValue("Row 4");

        // Create a font for the watermark
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create a text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,
            Opacity = 0.3f,
            IsBackground = true
        };

        // Freeze the first three rows (rows containing the watermark)
        sheet.FreezePanes("A4", 3, 0);

        // Set PDF save options with the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF with watermark and frozen rows
        workbook.Save("WatermarkedAndFrozen.pdf", options);
    }
}
