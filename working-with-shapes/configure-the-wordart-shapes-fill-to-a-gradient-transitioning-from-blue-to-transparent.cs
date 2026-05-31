using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class WordArtGradientFill
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // any preset style
            "Gradient WordArt",               // text
            2, 0,                             // upper left row, top offset
            2, 0,                             // upper left column, left offset
            100, 400);                        // height, width

        // Access the fill format of the WordArt shape
        FillFormat fillFormat = wordArt.Fill;

        // Set a two‑color gradient: opaque blue to fully transparent blue
        // GradientStyleType.Horizontal creates a left‑to‑right transition
        // Variant 1 corresponds to the first gradient variant in Excel
        fillFormat.SetTwoColorGradient(
            Color.Blue,      // first color (opaque)
            0.0,             // transparency for first color (0 = opaque)
            Color.Blue,      // second color (same hue)
            1.0,             // transparency for second color (1 = fully transparent)
            GradientStyleType.Horizontal,
            1);

        // Save the workbook
        workbook.Save("WordArtGradientFill.xlsx");
    }
}