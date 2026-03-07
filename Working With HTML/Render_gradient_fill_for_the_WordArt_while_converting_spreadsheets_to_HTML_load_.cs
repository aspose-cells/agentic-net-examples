using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape (any preset style works as a base)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Gradient WordArt",               // text
            2, 0,                            // upper left row, top offset
            2, 0,                            // upper left column, left offset
            100, 400);                       // height, width

        // Change the fill type to Gradient so we can configure it
        wordArt.Fill.FillType = FillType.Gradient;

        // Apply a two‑color gradient (light blue to dark blue, horizontal style)
        wordArt.Fill.SetTwoColorGradient(
            Color.LightBlue,                 // first color
            Color.DarkBlue,                  // second color
            GradientStyleType.Horizontal,    // gradient style
            1);                              // variant (1‑4)

        // Optionally fine‑tune the gradient direction/angle
        GradientFill gradientFill = wordArt.Fill.GradientFill;
        if (gradientFill != null)
        {
            // Linear gradient with a 45° angle, starting from the upper‑left corner
            gradientFill.SetGradient(GradientFillType.Linear, 45.0, GradientDirectionType.FromUpperLeftCorner);
        }

        // Save the workbook as HTML, preserving the WordArt with gradient fill
        workbook.Save("output.html", SaveFormat.Html);
    }
}