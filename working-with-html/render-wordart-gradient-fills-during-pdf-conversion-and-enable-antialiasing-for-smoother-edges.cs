// Title: C# – Export Gradient‑Filled WordArt to PDF with Anti‑Aliasing using Aspose.Cells
// Description: Shows how to insert a WordArt shape with the preset gradient style (WordArtStyle7), customize its font, and save the worksheet as a PDF. Aspose.Cells automatically renders the gradient fill and applies anti‑aliasing for smooth edges.
// Keywords: Aspose.Cells | C# | WordArt | gradient fill | PDF export | anti‑aliasing | PdfSaveOptions | shape rendering | preset WordArtStyle7 | Aspose.Cells example
// Common Searches: Aspose.Cells gradient WordArt PDF | C# export WordArt with gradient to PDF | enable anti‑aliasing in Aspose.Cells PDF | preserve WordArt styling when saving as PDF | add WordArt shape in Aspose.Cells
// Developer Intent: Generate a PDF that retains a gradient‑filled WordArt shape with smooth, anti‑aliased edges.
// Use Cases: Create marketing flyers or brochures with stylized WordArt headings and export them to high‑quality PDF. | Automate report generation that includes branded WordArt titles alongside data tables. | Produce print‑ready PDFs where shape rendering quality (gradient and anti‑aliasing) is critical.
// AI Prompts: Show C# code to add a gradient WordArt shape and export the workbook to PDF using Aspose.Cells. | How can I enable anti‑aliasing for shapes when saving a workbook as PDF with Aspose.Cells? | Provide an example that sets PdfSaveOptions.AntiAliasing = true and adds a WordArtStyle7 shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a WordArt shape with the preset gradient style (WordArtStyle7), customize its font, and save the worksheet as a PDF. Aspose.Cells automatically renders the gradient fill and applies anti‑aliasing for smooth edges.
class WordArtPdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape with a gradient preset style (WordArtStyle7)
            ShapeCollection shapes = sheet.Shapes;
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,   // gradient fill style
                "Gradient WordArt",                 // text
                2, 0,                               // top row and vertical offset (pixels)
                2, 0,                               // left column and horizontal offset (pixels)
                100, 400);                          // height and width (pixels)

            // Optional: customize the text appearance
            if (wordArt.IsWordArt)
            {
                TextEffectFormat textEffect = wordArt.TextEffect;
                textEffect.FontName = "Arial";
                textEffect.FontSize = 36;
                textEffect.FontBold = true;
            }

            // Save the workbook as PDF (default rendering options are sufficient)
            workbook.Save("WordArtGradient.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
