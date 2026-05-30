using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyShadowToWordArt
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that already contains a gradient fill
        // WordArtStyle6 = Gradient Fill - Gray
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6,
            "Gradient WordArt",
            5,      // topRow
            0,      // top (pixel offset)
            5,      // leftColumn
            0,      // left (pixel offset)
            100,    // height (pixels)
            400);   // width (pixels)

        // OPTIONAL: Customize the gradient fill further (preserve the gradient)
        // Ensure the fill type is gradient
        wordArt.Fill.FillType = FillType.Gradient;
        GradientFill gradient = wordArt.Fill.GradientFill;
        if (gradient != null)
        {
            // Define a two‑color diagonal gradient (light gray to dark gray)
            gradient.SetTwoColorGradient(
                Color.LightGray,          // first color
                Color.DarkGray,           // second color
                GradientStyleType.DiagonalDown,
                1);                       // variant
        }

        // Apply a shadow effect to the WordArt
        ShadowEffect shadow = wordArt.ShadowEffect;

        // Set shadow color (black)
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Color = Color.Black;
        shadow.Color = shadowColor;

        // You can also adjust other shadow properties if needed, e.g., transparency, blur, offset, etc.
        // shadow.Transparency = 0.5; // 50% transparent (example)

        // Save the workbook as PDF – the shadow and gradient will be preserved in the PDF output
        workbook.Save("WordArtWithShadow.pdf", SaveFormat.Pdf);
    }
}