// Title: Create a WordArt‑style text watermark and export an Excel workbook to PDF using Aspose.Cells for .NET
// Description: This C# sample builds a new Workbook, adds sample cells, defines a Calibri 68‑point bold‑italic font, and creates a RenderingWatermark with the text “CONFIDENTIAL”. The overlay is centered, rotated 45°, set to 30 % opacity, scaled to 75 % of the page, and placed behind the worksheet content. The watermark is attached to PdfSaveOptions and the workbook is saved as a PDF that shows the overlay.
// Keywords: Aspose.Cells PDF watermark | C# RenderingWatermark | Excel to PDF text overlay | WordArt watermark Aspose | PdfSaveOptions Watermark | Aspose.Cells .NET example | export Excel as PDF with watermark | semi‑transparent watermark C#
// Common Searches: How to add a rotated text watermark when saving Excel to PDF with Aspose.Cells | C# code for WordArt‑style watermark in PDF export | Set background watermark behind worksheet data using Aspose.Cells | Apply partially transparent watermark to PDF generated from a workbook | Export Excel file to PDF with custom font watermark in .NET
// Developer Intent: Embed a custom WordArt‑style text overlay into the PDF produced from an Aspose.Cells workbook.
// Use Cases: Confidential reports where the overlay must appear behind all cells | Brand‑marked PDFs with a slanted company name or slogan | Draft copies labeled with a visible “DRAFT” stamp for internal review | Legal documents that require a “CONFIDENTIAL” notice on every page | Marketing decks that need a subtle background tagline
// AI Prompts: Modify the example to use a red font color and 0.5 opacity for the watermark. | Generate code that applies different watermarks to odd and even pages during PDF export. | Show how to replace the text watermark with an image while preserving rotation and scaling. | Explain how to adjust the watermark to appear in front of worksheet content. | Provide a version that reads the watermark text from a cell value.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkPdf
{
    // This C# sample builds a new Workbook, adds sample cells, defines a Calibri 68‑point bold‑italic font, and creates a RenderingWatermark with the text “CONFIDENTIAL”. The overlay is centered, rotated 45°, set to 30 % opacity, scaled to 75 % of the page, and placed behind the worksheet content. The watermark is attached to PdfSaveOptions and the workbook is saved as a PDF that shows the overlay.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to make the watermark visible over content
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue("More Data");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Create a rendering font for the WordArt style watermark
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for a typical WordArt effect
                Rotation = 45,
                // Make it semi‑transparent so underlying content remains readable
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place the watermark behind the worksheet content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            workbook.Save("output_watermark.pdf", pdfOptions);
        }
    }
}
