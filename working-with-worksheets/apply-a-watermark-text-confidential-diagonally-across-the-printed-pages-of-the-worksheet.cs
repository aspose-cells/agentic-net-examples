// Title: C# – Add a Diagonal CONFIDENTIAL Text Watermark When Saving a Worksheet to PDF with Aspose.Cells
// Description: Creates a new workbook, defines a RenderingWatermark with the text "CONFIDENTIAL", sets a 45° rotation, 30% opacity, background placement, centered alignment, and full‑page scaling, then assigns the watermark to PdfSaveOptions and saves the worksheet as a printable PDF.
// Keywords: Aspose.Cells C# watermark PDF | RenderingWatermark rotation | diagonal text watermark Aspose | PdfSaveOptions watermark | confidential watermark worksheet | Aspose.Cells print watermark
// Common Searches: how to add a diagonal watermark to a worksheet PDF using Aspose.Cells | Aspose.Cells set watermark opacity and rotation | C# export worksheet to PDF with background text watermark | Aspose.Cells RenderingWatermark example
// Developer Intent: Apply a diagonal CONFIDENTIAL text watermark to the printed pages of a worksheet during PDF export.
// Use Cases: Produce legal or internal reports that display a semi‑transparent confidential label on every printed page. | Brand printable worksheets with a company logo or slogan across the page diagonal. | Automate batch generation of PDFs that require a consistent watermark for compliance.
// AI Prompts: Generate code to change the watermark text, font size, and color based on worksheet cell values in Aspose.Cells. | Show how to assign a different watermark to each sheet when exporting a multi‑sheet workbook to PDF. | Explain how to add a watermark only to the PDF output without affecting the workbook view in the Aspose.Cells UI.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a new workbook, defines a RenderingWatermark with the text "CONFIDENTIAL", sets a 45° rotation, 30% opacity, background placement, centered alignment, and full‑page scaling, then assigns the watermark to PdfSaveOptions and saves the worksheet as a printable PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to have visible content)
        worksheet.Cells["A1"].PutValue("Sample data for printing with watermark.");

        // Create a rendering font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create the text watermark with the desired text and font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Position the watermark diagonally across the page
            Rotation = 45f,
            // Make it semi‑transparent
            Opacity = 0.3f,
            // Place it behind the worksheet content
            IsBackground = true,
            // Center it horizontally and vertically
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            // Scale to fit the page
            ScaleToPagePercent = 100
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF (the PDF can be printed with the watermark)
        workbook.Save("WorksheetWithWatermark.pdf", saveOptions);
    }
}
