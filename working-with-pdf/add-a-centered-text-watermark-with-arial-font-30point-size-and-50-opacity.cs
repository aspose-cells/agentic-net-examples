// Title: Add a Centered Arial 30‑pt Text Watermark (50% Opacity) to PDF with Aspose.Cells for .NET
// Description: This C# example creates a new Workbook, inserts sample data, defines an Arial 30‑point RenderingFont, and applies a semi‑transparent watermark using RenderingWatermark. The watermark is horizontally and vertically centered, set to 50 % opacity, placed behind the page content, and saved through PdfSaveOptions as a PDF file.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark opacity | Arial 30‑point watermark | centered text watermark Aspose | PDF generation from Excel | Aspose.Cells .NET example
// Common Searches: aspnet add centered watermark to PDF using Aspose.Cells | how to set watermark opacity in Aspose.Cells PDF output | C# RenderingWatermark Arial font example | create semi transparent text watermark in PDF with Aspose
// Developer Intent: Generate a PDF from an Excel workbook that includes a centered, semi‑transparent text overlay.
// Use Cases: Flag confidential Excel reports as PDFs with a light‑gray overlay before sharing. | Apply corporate branding or legal notices across all exported PDFs. | Produce draft PDFs that display a non‑intrusive watermark indicating they are not final.
// AI Prompts: Write C# code using Aspose.Cells to place a centered Arial watermark at 30 pt with 50 % opacity in a PDF. | Show how to modify the watermark text, font size, color, or opacity in the given Aspose.Cells snippet. | Explain steps to replace the text watermark with an image while keeping it centered and semi‑transparent in PDF output.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example creates a new Workbook, inserts sample data, defines an Arial 30‑point RenderingFont, and applies a semi‑transparent watermark using RenderingWatermark. The watermark is horizontally and vertically centered, set to 50 % opacity, placed behind the page content, and saved through PdfSaveOptions as a PDF file.
class AddWatermark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample content to demonstrate the watermark effect
        sheet.Cells["A1"].PutValue("Sample data for the worksheet");

        // Define the font for the watermark: Arial, 30‑point size
        RenderingFont font = new RenderingFont("Arial", 30)
        {
            // Optional: set a light color for better visibility
            Color = Color.LightGray
        };

        // Create a centered text watermark with the specified font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,   // Horizontal center
            VAlignment = TextAlignmentType.Center,   // Vertical center
            Opacity = 0.5f,                          // 50% opacity
            IsBackground = true                      // Place behind page contents
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("WatermarkedOutput.pdf", options);
    }
}
