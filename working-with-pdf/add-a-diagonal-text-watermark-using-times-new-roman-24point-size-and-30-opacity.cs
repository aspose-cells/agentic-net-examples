// Title: Add a Diagonal Text Watermark to PDF with Aspose.Cells for .NET (C#) – Times New Roman 24 pt, 30 % Opacity
// Description: This C# example creates a Workbook, defines a Times New Roman 24‑point RenderingFont, and applies a 45° rotated, 30 % opaque, centered text watermark behind the worksheet content using RenderingWatermark and PdfSaveOptions, then saves the file as DiagonalWatermark.pdf.
// Keywords: Aspose.Cells PDF watermark C# | diagonal text watermark Aspose.Cells | RenderingWatermark opacity | Times New Roman watermark | PdfSaveOptions watermark | C# add watermark to PDF | Aspose.Cells rendering font
// Common Searches: C# add diagonal watermark to PDF with Aspose.Cells | How to set watermark opacity in Aspose.Cells PDF export | Render text watermark behind worksheet content Aspose.Cells | Aspose.Cells PDFSaveOptions watermark rotation | Create confidential PDF with Aspose.Cells
// Developer Intent: Generate a PDF from a workbook that includes a centered, diagonal text watermark (custom font, size, opacity) placed behind the sheet data.
// Use Cases: Mark confidential reports with a semi‑transparent diagonal label. | Brand exported spreadsheets with a company logo or slogan across each PDF page. | Add regulatory disclaimer watermarks to PDFs without hiding cell values. | Create draft versions of documents where the watermark indicates 'DRAFT' across pages.
// AI Prompts: Generate code to change the watermark text, font family, size, rotation, and opacity based on user parameters in Aspose.Cells. | Show how to apply different watermarks to individual worksheets when exporting to PDF. | Explain how to add an image watermark together with a text watermark using Aspose.Cells. | Provide a step‑by‑step guide for removing a watermark from an existing PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example creates a Workbook, defines a Times New Roman 24‑point RenderingFont, and applies a 45° rotated, 30 % opaque, centered text watermark behind the worksheet content using RenderingWatermark and PdfSaveOptions, then saves the file as DiagonalWatermark.pdf.
class AddDiagonalWatermark
{
    static void Main()
    {
        // Create a new workbook and add some sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content");

        // Create a rendering font: Times New Roman, 24‑point size
        RenderingFont font = new RenderingFont("Times New Roman", 24);

        // Initialize a text watermark with the specified font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Rotate 45 degrees for a diagonal appearance
            Rotation = 45f,
            // Set opacity to 30%
            Opacity = 0.3f,
            // Place the watermark behind the worksheet content
            IsBackground = true,
            // Center the watermark on the page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the diagonal watermark
        workbook.Save("DiagonalWatermark.pdf", options);
    }
}
