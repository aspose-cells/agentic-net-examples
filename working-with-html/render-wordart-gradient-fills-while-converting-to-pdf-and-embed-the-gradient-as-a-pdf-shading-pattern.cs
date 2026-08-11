// Title: C# – Add Gradient WordArt and Export to PDF with Aspose.Cells
// Description: Creates a workbook, inserts a WordArt shape using a preset gradient style, customizes a two‑color horizontal gradient, and saves the file as PDF where the gradient is embedded as a PDF shading pattern.
// Keywords: Aspose.Cells | C# | WordArt gradient | PDF shading pattern | export to PDF | gradient fill | preset WordArt style | two‑color gradient
// Common Searches: Aspose.Cells add WordArt gradient C# | export WordArt with gradient to PDF | PDF shading pattern for Excel shapes | set custom gradient for WordArt in .NET | preserve gradient fill when converting Excel to PDF
// Developer Intent: Generate a WordArt shape with a gradient fill and ensure the gradient is retained as a shading pattern in the resulting PDF.
// Use Cases: Design marketing flyers with gradient WordArt headings and produce print‑ready PDFs. | Automate report creation where section titles use gradient WordArt for visual impact. | Create branded PDFs that embed corporate‑color gradient WordArt for consistent styling.
// AI Prompts: Show how to change the gradient colors and direction of a WordArt shape before saving to PDF with Aspose.Cells. | Provide example code to add multiple WordArt shapes, each with a different gradient, and keep the gradients in the PDF output. | Explain how to verify that a generated PDF contains shading patterns for WordArt gradients using a PDF inspection tool.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, inserts a WordArt shape using a preset gradient style, customizes a two‑color horizontal gradient, and saves the file as PDF where the gradient is embedded as a PDF shading pattern.
class WordArtGradientPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (WordArtStyle7: Gradient Fill - Blue, Accent 1, Reflection)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,   // upperLeftRow, top
            2, 0,   // upperLeftColumn, left
            200,    // height
            400);   // width

        // Ensure the fill type is set to Gradient (preset already does this, but set explicitly for clarity)
        wordArt.Fill.FillType = FillType.Gradient;

        // Optionally customize the gradient further using FillFormat
        // Here we define a simple two‑color horizontal gradient from Blue to LightBlue
        wordArt.Fill.SetTwoColorGradient(
            Color.Blue,          // first color
            Color.LightBlue,     // second color
            GradientStyleType.Horizontal,
            1);                  // variant (1‑4)

        // Save the workbook as PDF.
        // Aspose.Cells automatically embeds the shape's gradient as a PDF shading pattern.
        workbook.Save("WordArtGradient.pdf", SaveFormat.Pdf);
    }
}
