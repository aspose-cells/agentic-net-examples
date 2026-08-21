// Title: C# – Render Gradient WordArt with Anti‑Aliasing When Converting Excel to PDF using Aspose.Cells
// Description: This example creates a workbook, inserts a WordArt shape using the preset WordArtStyle7 gradient, customizes its font, and saves the sheet as a PDF. PdfSaveOptions applies built‑in anti‑aliasing so the gradient fill and edges appear smooth in the final document.
// Keywords: Aspose.Cells | C# | WordArt gradient | Excel to PDF | PdfSaveOptions | anti‑aliasing | shape rendering | preset WordArtStyle7 | GitHub sample | .NET PDF conversion
// Common Searches: How to preserve WordArt gradient when exporting Excel to PDF with Aspose.Cells | Enable anti‑aliasing for shapes in PDF output using Aspose.Cells .NET | Add preset gradient WordArt to a worksheet and convert to high‑quality PDF | PdfSaveOptions settings for smoother shape rendering in Aspose.Cells
// Developer Intent: Export an Excel worksheet that contains a gradient‑filled WordArt shape to PDF while ensuring the graphics are anti‑aliased for crisp visual quality.
// Use Cases: Create marketing flyers in Excel with gradient WordArt titles and generate print‑ready PDFs. | Automate report generation that includes stylized WordArt headings, delivering polished PDFs to stakeholders. | Batch‑process workbooks to insert gradient WordArt labels and export each file as a high‑quality PDF.
// AI Prompts: Generate C# code that adds a WordArt shape with a preset gradient fill to an Aspose.Cells worksheet and saves it as a PDF with anti‑aliasing. | Explain the impact of PdfSaveOptions on the rendering quality of WordArt and other shapes during Excel‑to‑PDF conversion in Aspose.Cells. | Outline steps to verify that a gradient WordArt shape appears correctly and smoothly in the exported PDF.

using System;
using System.Drawing;
using System.Drawing.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, inserts a WordArt shape using the preset WordArtStyle7 gradient, customizes its font, and saves the sheet as a PDF. PdfSaveOptions applies built‑in anti‑aliasing so the gradient fill and edges appear smooth in the final document.
class RenderWordArtPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape with a gradient preset style (WordArtStyle7)
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Gradient WordArt",               // Text
                2, 0,                             // topRow, top (pixel offset)
                2, 0,                             // leftColumn, left (pixel offset)
                100, 400);                        // height, width (pixel)

            // Customize the text appearance if the shape is WordArt
            if (wordArt.IsWordArt)
            {
                TextEffectFormat textEffect = wordArt.TextEffect;
                textEffect.FontBold = true;
                textEffect.FontSize = 24;
                textEffect.FontName = "Arial";
            }

            // Configure PDF save options (default anti‑aliasing is applied by Aspose.Cells)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            string outputPath = "WordArtGradient.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
