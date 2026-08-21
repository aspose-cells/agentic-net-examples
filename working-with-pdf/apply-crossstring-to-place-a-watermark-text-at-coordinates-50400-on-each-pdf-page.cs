// Title: Add a text watermark at (50,400) to every page of a PDF using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a RenderingFont, builds a RenderingWatermark with the text "CONFIDENTIAL" positioned 50 pts from the left and 400 pts from the top, assigns it to PdfSaveOptions, and saves the workbook as a PDF where the watermark is rendered on each page.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark offset | PdfSaveOptions watermark | watermark coordinates Aspose.Cells | cross‑string watermark | C# PDF export watermark | Aspose.Cells RenderingFont
// Common Searches: Aspose.Cells place text watermark at specific coordinates | C# add watermark to every page of PDF with Aspose.Cells | set RenderingWatermark offset X Y Aspose.Cells | PdfSaveOptions watermark background PDF export | Aspose.Cells cross‑string watermark example
// Developer Intent: Add a semi‑transparent text watermark at the (50,400) point on each page of a PDF generated from an Excel workbook.
// Use Cases: Generate confidential reports with a fixed‑position watermark on all pages. | Embed a company disclaimer at a precise location when exporting Excel templates to PDF. | Automate batch PDF creation from multiple workbooks while maintaining consistent watermark placement for compliance.
// AI Prompts: Show how to rotate the watermark 45° while keeping the (50,400) offset in Aspose.Cells. | Provide code that reads the watermark text from a variable and applies it using RenderingWatermark. | Explain how to assign different watermarks to odd and even pages when saving a workbook as PDF.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, defines a RenderingFont, builds a RenderingWatermark with the text "CONFIDENTIAL" positioned 50 pts from the left and 400 pts from the top, assigns it to PdfSaveOptions, and saves the workbook as a PDF where the watermark is rendered on each page.
class WatermarkExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional)
        worksheet.Cells["A1"].PutValue("Sample data for PDF with watermark");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36)
        {
            Bold = true,
            Color = Color.Gray
        };

        // Create a text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Position the watermark at (50, 400) points from the top‑left corner
            HAlignment = TextAlignmentType.Left,
            VAlignment = TextAlignmentType.Top,
            OffsetX = 50f,
            OffsetY = 400f,

            // Additional appearance settings
            Rotation = 0f,
            Opacity = 0.3f,
            IsBackground = true
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF; the watermark will appear on every page
        workbook.Save("WatermarkedOutput.pdf", pdfOptions);
    }
}
