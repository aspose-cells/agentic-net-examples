// Title: Add a diagonal CONFIDENTIAL watermark when saving an Excel workbook to PDF with Aspose.Cells for .NET
// Description: Loads an .xlsx file, creates a RenderingWatermark with centered alignment, 45° rotation, 30% opacity and 75% page scaling, attaches it to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears behind the content on every page.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF | PDF watermark | RenderingWatermark | PdfSaveOptions | diagonal watermark | confidential watermark | background watermark | watermark opacity | watermark scaling
// Common Searches: Aspose.Cells add diagonal watermark to PDF | C# save Excel as PDF with watermark | RenderingWatermark rotation opacity example | PdfSaveOptions watermark settings Aspose.Cells | how to put CONFIDENTIAL text on each PDF page from Excel
// Developer Intent: Generate a PDF from an Excel workbook and embed a semi‑transparent diagonal text watermark on every page.
// Use Cases: Distribute confidential reports with a 'CONFIDENTIAL' label behind the data. | Brand exported PDFs with a company slogan or logo as a background watermark. | Apply legal or compliance notices to spreadsheet PDFs before sharing.
// AI Prompts: Write C# code using Aspose.Cells to load an .xlsx file and save it as a PDF with a 45° rotated, 30% opaque 'CONFIDENTIAL' watermark behind the content. | Explain how to customize font, rotation, opacity, and scaling of a RenderingWatermark in PdfSaveOptions for Excel‑to‑PDF conversion. | Provide step‑by‑step instructions for adding a background watermark to all pages of a PDF generated from a workbook with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    // Loads an .xlsx file, creates a RenderingWatermark with centered alignment, 45° rotation, 30% opacity and 75% page scaling, attaches it to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears behind the content on every page.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourceFile);

            // Create a font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
            {
                Bold = true,
                Italic = true,
                Color = Color.Blue
            };

            // Create a text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,

                // Rotate the watermark for a diagonal appearance
                Rotation = 45,

                // Set opacity (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.3f,

                // Scale the watermark relative to the page size (percentage)
                ScaleToPagePercent = 75,

                // Place the watermark behind the page content
                IsBackground = true
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied to each page
            string outputFile = "output_watermark.pdf";
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook saved as PDF with watermark: {outputFile}");
        }
    }
}
