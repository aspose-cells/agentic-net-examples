// Title: C# – Set Worksheet Zoom to 110 % and Export PDF with Embedded Fonts using Aspose.Cells
// Description: Creates a workbook, applies a 110 % percent‑scale zoom via PageSetup.Zoom and IsPercentScale, configures PdfSaveOptions to embed standard Windows TrueType fonts with Identity encoding and a default Arial font, then saves the sheet as a PDF.
// Keywords: Aspose.Cells | C# | .NET | worksheet zoom | 110% zoom | PageSetup.Zoom | IsPercentScale | PDF export | embed fonts | PdfSaveOptions | EmbedStandardWindowsFonts | Identity encoding | default font | cross‑platform PDF
// Common Searches: Aspose.Cells set worksheet zoom before PDF export | How to embed TrueType fonts in PDF using Aspose.Cells .NET | 110 percent zoom PDF Aspose.Cells example | PageSetup.IsPercentScale effect on PDF scaling | C# code to save workbook as PDF with embedded fonts
// Developer Intent: Apply a 110 % zoom to a worksheet and generate a PDF that includes embedded fonts.
// Use Cases: Produce printable PDFs where the content is enlarged for readability while guaranteeing font fidelity on any device. | Create standardized reports that require percent‑based scaling and must retain exact typography across platforms. | Export multilingual worksheets to PDF with full Unicode support and embedded Windows fonts to prevent missing glyphs.
// AI Prompts: Show C# code that sets PageSetup.Zoom to 110% and saves the workbook as a PDF with embedded fonts using Aspose.Cells. | Explain how to configure PdfSaveOptions for font embedding and Identity encoding in Aspose.Cells .NET. | What does PageSetup.IsPercentScale do when exporting a worksheet to PDF with Aspose.Cells?

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsZoomPdfDemo
{
    // Creates a workbook, applies a 110 % percent‑scale zoom via PageSetup.Zoom and IsPercentScale, configures PdfSaveOptions to embed standard Windows TrueType fonts with Identity encoding and a default Arial font, then saves the sheet as a PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content in PDF)
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["A2"].PutValue("Zoom set to 110% and fonts embedded in PDF.");

            // Set the zoom level to 110%
            // Use PageSetup.Zoom for printing/scaling purposes
            worksheet.PageSetup.Zoom = 110;
            // Ensure the scaling mode is percent based
            worksheet.PageSetup.IsPercentScale = true;

            // Configure PDF save options to embed fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed standard Windows TrueType fonts (required for embedding)
                EmbedStandardWindowsFonts = true,
                // Use identity encoding for full Unicode support
                FontEncoding = PdfFontEncoding.Identity,
                // Optionally set a default font in case some cells lack explicit font
                DefaultFont = "Arial"
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("ZoomedWithEmbeddedFonts.pdf", pdfOptions);
        }
    }
}
