// Title: Apply a Shadow to Gradient‑Filled WordArt and Export to PDF with Aspose.Cells for .NET
// Description: Creates a workbook, adds a WordArt shape using PresetWordArtStyle.WordArtStyle6, sets a two‑color diagonal gradient, applies a semi‑transparent black shadow via the ShadowEffect property, and saves the sheet as a PDF while preserving both visual effects.
// Keywords: Aspose.Cells | C# | .NET | WordArt gradient | ShadowEffect | PDF export | PresetWordArtStyle | GradientFill | CellsColor | shadow on WordArt
// Common Searches: Aspose.Cells add shadow to WordArt | WordArt gradient PDF export C# | preserve WordArt effects when saving as PDF | how to set shadow color and transparency in Aspose.Cells | apply diagonal gradient to WordArt in .NET
// Developer Intent: Add a gradient‑filled WordArt shape, apply a shadow effect, and generate a PDF that retains both styles using Aspose.Cells for .NET.
// Use Cases: Design marketing flyers with stylized headings that keep gradient and shadow in printable PDFs. | Automate creation of certificates where the title uses gradient WordArt with a shadow for a professional look. | Produce data dashboards that export to PDF, maintaining visual consistency of WordArt titles.
// AI Prompts: Show how to modify the shadow offset, blur radius, and opacity for WordArt in the Aspose.Cells example. | Generate code that adds multiple WordArt objects, each with distinct gradient colors and shadow settings, before PDF conversion. | Explain how to change the gradient direction and colors while preserving the shadow effect in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a WordArt shape using PresetWordArtStyle.WordArtStyle6, sets a two‑color diagonal gradient, applies a semi‑transparent black shadow via the ShadowEffect property, and saves the sheet as a PDF while preserving both visual effects.
class WordArtShadowPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (WordArtStyle6)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6,
            "Gradient WordArt",
            5,   // topRow
            0,   // top (pixels)
            5,   // leftColumn
            0,   // left (pixels)
            100, // height (pixels)
            400  // width (pixels)
        );

        // Ensure the fill type is gradient
        wordArt.Fill.FillType = FillType.Gradient;

        // Configure a custom two‑color gradient (e.g., LightGray to DarkGray)
        GradientFill gradient = wordArt.Fill.GradientFill;
        gradient.SetTwoColorGradient(
            Color.LightGray,   // first color
            Color.DarkGray,    // second color
            GradientStyleType.DiagonalDown,
            1                  // variant
        );

        // Apply a shadow effect to the WordArt
        ShadowEffect shadow = wordArt.ShadowEffect;

        // Create a CellsColor for the shadow (e.g., semi‑transparent black)
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Color = Color.FromArgb(128, 0, 0, 0); // 50% transparent black
        shadow.Color = shadowColor;

        // Save the workbook as PDF (shadow and gradient are preserved)
        workbook.Save("WordArtWithShadow.pdf", SaveFormat.Pdf);
    }
}
