// Title: Add a Shadow to WordArt with Gradient Fill and Export as PDF using Aspose.Cells for .NET
// Description: Demonstrates how to insert a WordArt shape with a preset gradient style, apply a black shadow via the ShadowEffect object, keep the original gradient intact, and save the workbook as a PDF.
// Keywords: Aspose.Cells WordArt shadow | gradient fill WordArt PDF | C# add shadow to WordArt | preserve WordArt gradient Aspose | export Excel to PDF with shadow
// Common Searches: how to add shadow to WordArt in Aspose.Cells | keep gradient fill when applying WordArt shadow .NET | C# Aspose.Cells shadow effect for WordArt | save WordArt with shadow as PDF using Aspose
// Developer Intent: Apply a shadow to a WordArt shape while retaining its gradient fill, then generate a PDF document.
// Use Cases: Design report headings with gradient WordArt and subtle shadows for clearer visual hierarchy. | Programmatically create marketing flyers that need WordArt depth without losing color transitions. | Batch‑process Excel templates containing WordArt, ensuring each shape keeps its gradient and receives a uniform shadow in the final PDFs.
// AI Prompts: Generate C# code to set shadow offset, blur radius, and opacity for a WordArt shape in Aspose.Cells while preserving its gradient. | Show how to apply both outer and inner shadow types to multiple WordArt objects before exporting to PDF. | Explain how to read the existing gradient stops of a WordArt shape after a shadow is added and modify the colors programmatically.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to insert a WordArt shape with a preset gradient style, apply a black shadow via the ShadowEffect object, keep the original gradient intact, and save the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add WordArt with a preset style that already contains a gradient fill (WordArtStyle7)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Shadowed WordArt",
            5,   // topRow
            0,   // top (pixel offset)
            5,   // leftColumn
            0,   // left (pixel offset)
            150, // height (pixel)
            400  // width (pixel)
        );

        // Apply a shadow effect while preserving the existing gradient fill
        ShadowEffect shadowEffect = wordArt.ShadowEffect;

        // Set shadow color (e.g., black)
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Color = Color.Black;
        shadowEffect.Color = shadowColor;

        // Save the workbook as PDF
        workbook.Save("WordArtShadow.pdf", SaveFormat.Pdf);
    }
}
