// Title: Apply a Light Gray‑to‑Dark Blue Horizontal Gradient to a WordArt Watermark in Aspose.Cells (C#)
// Description: Creates a new workbook, adds a WordArt shape with the text "CONFIDENTIAL", switches its fill to a horizontal gradient (light gray → dark blue), configures a semi‑transparent RenderingWatermark, and saves the result as a PDF. Demonstrates gradient styling of WordArt for watermarking in Aspose.Cells for .NET.
// Keywords: Aspose.Cells gradient WordArt | C# WordArt watermark | horizontal gradient fill Aspose.Cells | two‑color gradient WordArt PDF | Aspose.Cells rendering watermark | gradient fill shape Aspose.Cells | PDF export with WordArt watermark
// Common Searches: Aspose.Cells set gradient fill on WordArt C# | how to create gradient WordArt watermark in .NET | horizontal light gray to dark blue gradient Aspose.Cells | export workbook to PDF with WordArt watermark | change WordArt fill type to gradient Aspose.Cells
// Developer Intent: Add a WordArt shape, apply a horizontal light‑gray‑to‑dark‑blue gradient, and generate a PDF where the shape serves as a semi‑transparent watermark.
// Use Cases: Brand‑consistent confidential PDFs with a corporate‑color gradient watermark. | Automated report pipelines that mark drafts or final versions using gradient WordArt watermarks. | Visually distinct security labels across multiple worksheets in a spreadsheet export.
// AI Prompts: Generate C# code to change the gradient direction to vertical and use red and orange colors for the WordArt watermark. | Provide an example that adds a different gradient WordArt watermark to each worksheet in a single workbook. | Explain how to modify the RenderingWatermark opacity while keeping the WordArt gradient unchanged.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a new workbook, adds a WordArt shape with the text "CONFIDENTIAL", switches its fill to a horizontal gradient (light gray → dark blue), configures a semi‑transparent RenderingWatermark, and saves the result as a PDF. Demonstrates gradient styling of WordArt for watermarking in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape that will serve as the visual watermark
        // Parameters: preset style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // Base style (gradient fill - gray)
            "CONFIDENTIAL",                  // Watermark text
            5, 0,                            // Row and top offset
            5, 0,                            // Column and left offset
            300, 100);                       // Height and width

        // Change the fill type to gradient so we can customize it
        wordArt.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object and set a two‑color gradient
        // Light gray -> Dark blue, horizontal direction, first variant
        GradientFill gradientFill = wordArt.Fill.GradientFill;
        gradientFill.SetTwoColorGradient(
            Color.LightGray,   // First color (light gray)
            Color.DarkBlue,    // Second color (dark blue)
            GradientStyleType.Horizontal,
            1);                // Variant

        // Create a RenderingFont for the text watermark (used by PDF rendering)
        RenderingFont renderingFont = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.Black
        };

        // Create the RenderingWatermark with the same text
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", renderingFont)
        {
            Opacity = 0.3f,                     // Semi‑transparent
            Rotation = 45,                      // Diagonal appearance
            IsBackground = true,                // Render behind page content
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center
        };

        // Save the workbook as PDF with the watermark applied
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };
        workbook.Save("WatermarkedGradientWordArt.pdf", saveOptions);
    }
}
