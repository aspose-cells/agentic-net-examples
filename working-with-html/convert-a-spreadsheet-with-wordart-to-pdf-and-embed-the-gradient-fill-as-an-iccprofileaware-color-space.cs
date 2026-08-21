// Title: Export WordArt with Gradient Fill to PDF/A‑1b and Embed ICC‑Aware Color Profile using Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, add a WordArt shape with a vertical two‑color gradient, and save it as a PDF/A‑1b file. The PdfSaveOptions embed standard Windows fonts and include the ICC color profile so the gradient renders accurately on any device.
// Keywords: Aspose.Cells C# | WordArt gradient | PDF export | PDF/A-1b | ICC color profile | EmbedStandardWindowsFonts | FillFormat gradient | PresetWordArtStyle | gradient fill PDF | color fidelity Aspose.Cells
// Common Searches: Aspose.Cells add WordArt with gradient and export to PDF | how to embed ICC profile when saving PDF with Aspose.Cells | save workbook as PDF/A-1b with gradient colors | C# Aspose.Cells gradient fill WordArt PDF | preserve gradient colors in PDF using Aspose.Cells
// Developer Intent: Generate a PDF/A‑1b document that contains a WordArt heading with a vertical gradient, ensuring the gradient colors are stored in an ICC‑aware color space.
// Use Cases: Create marketing brochures where WordArt headings keep exact brand gradients in archived PDFs. | Produce compliance‑ready reports (PDF/A‑1b) that retain color fidelity for decorative text. | Export spreadsheets with decorative WordArt to PDF while guaranteeing consistent appearance across printers and screens.
// AI Prompts: Show C# code to apply a custom two‑color gradient to a WordArt shape before saving as PDF with Aspose.Cells. | Explain how PdfSaveOptions can embed an ICC profile for gradient fills in Aspose.Cells. | Provide an example that converts a workbook containing multiple WordArt objects with different gradients to a PDF/A‑2b file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWordArtToPdf
{
    // Demonstrates how to create a workbook, add a WordArt shape with a vertical two‑color gradient, and save it as a PDF/A‑1b file. The PdfSaveOptions embed standard Windows fonts and include the ICC color profile so the gradient renders accurately on any device.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Add a WordArt shape with a preset style that already contains a
            // gradient fill (WordArtStyle7 – Gradient Fill - Blue, Accent 1,
            // Reflection). The parameters are: style, text, topRow, top,
            // leftColumn, left, height, width.
            // -----------------------------------------------------------------
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,
                "Aspose.Cells WordArt",
                2,          // topRow
                10,         // top (pixels)
                2,          // leftColumn
                10,         // left (pixels)
                100,        // height (pixels)
                400);       // width (pixels)

            // -----------------------------------------------------------------
            // Customize the gradient fill of the WordArt.
            // First, set the fill type to Gradient, then obtain the FillFormat
            // and apply a two‑color gradient (e.g., from LightBlue to DarkBlue).
            // -----------------------------------------------------------------
            wordArt.Fill.FillType = FillType.Gradient;
            FillFormat fill = wordArt.Fill;
            fill.SetTwoColorGradient(
                Color.LightBlue,          // first gradient color
                Color.DarkBlue,           // second gradient color
                GradientStyleType.Vertical,
                1);                       // variant (1‑4)

            // -----------------------------------------------------------------
            // Prepare PDF save options.
            // EmbedStandardWindowsFonts ensures that the fonts used in the
            // WordArt are embedded in the PDF, which also embeds the color
            // profile information required for ICC‑aware rendering.
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedStandardWindowsFonts = true,   // embed fonts (helps with color profile)
                Compliance = PdfCompliance.PdfA1b   // optional: PDF/A‑1b compliance
            };

            // Save the workbook (including the WordArt) as a PDF file
            workbook.Save("WordArtWithGradient.pdf", pdfOptions);

            Console.WriteLine("PDF generated successfully with WordArt gradient fill.");
        }
    }
}
