// Title: Add a diagonal semi‑transparent watermark and freeze top rows in an Aspose.Cells worksheet (C#)
// Description: C# example that creates a workbook, freezes the first three rows, defines a light‑gray "CONFIDENTIAL" watermark with 45° rotation and 20% opacity using RenderingWatermark, and saves the sheet as a PDF with the watermark locked in view.
// Keywords: Aspose.Cells watermark C# | freeze panes Aspose.Cells | RenderingWatermark PDF export | diagonal watermark .NET | semi transparent watermark Aspose | freeze top rows worksheet | PDF save options Aspose.Cells | C# Aspose.Cells example
// Common Searches: how to add a diagonal watermark with Aspose.Cells | freeze first rows in Aspose.Cells worksheet | Aspose.Cells RenderingWatermark opacity rotation | export worksheet to PDF with background watermark C# | Aspose.Cells freeze panes and watermark together
// Developer Intent: Generate a PDF from a worksheet that includes a background watermark while keeping the header rows frozen.
// Use Cases: Confidential reports where the watermark must stay visible as users scroll. | Brand‑protected invoices with a logo watermark and locked title rows. | Multi‑page statements that need a repeated watermark and frozen header for consistent identification.
// AI Prompts: Show C# code that adds a RenderingWatermark with custom font, rotation, and opacity to PdfSaveOptions in Aspose.Cells. | Provide an Aspose.Cells example that freezes the top three rows before exporting the worksheet as a PDF with a diagonal watermark.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# example that creates a workbook, freezes the first three rows, defines a light‑gray "CONFIDENTIAL" watermark with 45° rotation and 20% opacity using RenderingWatermark, and saves the sheet as a PDF with the watermark locked in view.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Report";

        // Add sample data (the watermark will appear over these rows)
        worksheet.Cells["A1"].PutValue("Report Title");
        worksheet.Cells["A2"].PutValue("Generated on:");
        worksheet.Cells["B2"].PutValue(DateTime.Now);

        // Freeze the top three rows so the watermark area cannot be scrolled away
        // Freeze at cell A4 (row index 3) with 3 frozen rows and 0 frozen columns
        worksheet.FreezePanes("A4", 3, 0);

        // Create a font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            IsBackground = true,          // place behind page contents
            Rotation = 45,                // rotate 45 degrees
            Opacity = 0.2f,               // semi‑transparent
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center
        };

        // Set PDF save options to include the watermark
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("ReportWithWatermark.pdf", saveOptions);
    }
}
