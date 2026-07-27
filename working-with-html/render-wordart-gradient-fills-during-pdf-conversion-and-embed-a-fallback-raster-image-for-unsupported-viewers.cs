// Title: Render Gradient WordArt in PDF with a Raster Fallback using Aspose.Cells for .NET
// Description: This C# example shows how to add a WordArt shape with a two‑color gradient to an Excel workbook, configure a semi‑transparent PNG as a background RenderingWatermark, and save the workbook as a PDF. The raster watermark acts as a fallback for PDF viewers that cannot render gradient fills, while PdfSaveOptions ensure font and compatibility checks.
// Keywords: Aspose.Cells | C# | .NET | WordArt gradient | Excel to PDF conversion | RenderingWatermark | PDF fallback image | raster watermark | gradient fill support | PdfSaveOptions | cross‑platform PDF rendering
// Common Searches: Aspose.Cells gradient WordArt PDF export | add PNG watermark as fallback in PDF using Aspose.Cells | render WordArt gradient in PDF with Aspose.Cells .NET | PdfSaveOptions watermark raster image Aspose.Cells | Excel WordArt gradient not showing in PDF viewer
// Developer Intent: Create a gradient‑filled WordArt shape in an Excel workbook and export it to PDF while embedding a semi‑transparent raster watermark that serves as a visual fallback for PDF viewers lacking gradient support.
// Use Cases: Design marketing brochures where Excel WordArt provides branding, with a PNG fallback for legacy PDF readers. | Generate compliance‑oriented reports that need a visible watermark if gradient rendering fails. | Produce printable PDFs from Excel dashboards that retain visual fidelity across all viewer versions.
// AI Prompts: Generate C# code with Aspose.Cells to insert a WordArt shape using a custom two‑color gradient and export the sheet to PDF with a centered semi‑transparent PNG watermark as a background. | Explain how to set up PdfSaveOptions in Aspose.Cells to embed a raster image watermark that acts as a fallback for unsupported gradient fills. | Provide a step‑by‑step verification method to confirm that the gradient WordArt renders correctly in PDF and that the fallback watermark appears when the gradient is not supported.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWordArtPdfDemo
{
    // This C# example shows how to add a WordArt shape with a two‑color gradient to an Excel workbook, configure a semi‑transparent PNG as a background RenderingWatermark, and save the workbook as a PDF. The raster watermark acts as a fallback for PDF viewers that cannot render gradient fills, while PdfSaveOptions ensure font and compatibility checks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a WordArt shape with a preset gradient style (WordArtStyle7)
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,
                "Gradient WordArt",
                5,   // topRow
                10,  // top (pixels)
                2,   // leftColumn
                10,  // left (pixels)
                100, // height (pixels)
                400  // width (pixels)
            );

            // Ensure the fill type is gradient and customize the gradient colors
            wordArt.Fill.FillType = FillType.Gradient;
            GradientFill gradient = wordArt.Fill.GradientFill;
            if (gradient != null)
            {
                // Two‑color gradient: LightBlue to DarkBlue, horizontal style, variant 1
                gradient.SetTwoColorGradient(Color.LightBlue, Color.DarkBlue, GradientStyleType.Horizontal, 1);
            }

            // -----------------------------------------------------------------
            // Prepare a fallback raster image (e.g., PNG) to be used as a watermark.
            // This image will be visible in PDF viewers that cannot render the
            // WordArt gradient correctly.
            // -----------------------------------------------------------------
            string fallbackImagePath = "fallback.png"; // Path to your raster image
            if (!File.Exists(fallbackImagePath))
            {
                Console.WriteLine($"Fallback image not found: {fallbackImagePath}");
                return;
            }

            byte[] imageData = File.ReadAllBytes(fallbackImagePath);

            // Create a RenderingWatermark using the raster image data
            RenderingWatermark rasterWatermark = new RenderingWatermark(imageData)
            {
                // Position the watermark at the center of the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Make it semi‑transparent so the WordArt can be seen underneath
                Opacity = 0.4f,
                // Scale the image to 30% of the page size
                ScaleToPagePercent = 30,
                // Render it as a background element
                IsBackground = true
            };

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed the raster watermark as a fallback visual aid
                Watermark = rasterWatermark,
                // Ensure font fallback is attempted for any missing glyphs
                CheckWorkbookDefaultFont = true,
                // Keep default font compatibility checking enabled
                CheckFontCompatibility = true
            };

            // Save the workbook as PDF with the specified options
            string outputPdf = "WordArtWithFallback.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
        }
    }
}
