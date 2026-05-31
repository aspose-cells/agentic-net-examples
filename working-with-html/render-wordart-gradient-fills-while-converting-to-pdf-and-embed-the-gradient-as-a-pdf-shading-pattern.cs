using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class WordArtGradientPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,   // row, top offset
            2, 0,   // column, left offset
            100,    // height
            400);   // width

        // Ensure the fill type is set to Gradient to allow custom gradient settings
        wordArt.Fill.FillType = FillType.Gradient;

        // Apply a custom two‑color gradient to the WordArt fill
        // Gradient: Blue to LightBlue, horizontal direction, first variant
        wordArt.Fill.SetTwoColorGradient(
            Color.Blue,
            Color.LightBlue,
            GradientStyleType.Horizontal,
            1);

        // Optionally, adjust the gradient angle via the underlying GradientFill object
        GradientFill gradFill = wordArt.Fill.GradientFill;
        if (gradFill != null)
        {
            gradFill.Angle = 0; // horizontal
        }

        // Save the workbook as PDF; Aspose.Cells embeds the gradient as a PDF shading pattern
        workbook.Save("WordArtGradient.pdf", SaveFormat.Pdf);
    }
}