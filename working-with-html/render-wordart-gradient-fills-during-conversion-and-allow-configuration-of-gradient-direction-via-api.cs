using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class WordArtGradientDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            5,          // upperLeftRow
            0,          // top (pixels)
            5,          // upperLeftColumn
            0,          // left (pixels)
            200,        // height (pixels)
            100);       // width (pixels)

        // Ensure the fill type is set to Gradient to access GradientFill
        wordArt.Fill.FillType = FillType.Gradient;

        // Get the GradientFill object from the shape's fill
        GradientFill gradientFill = wordArt.Fill.GradientFill;

        // Optionally define the two colors of the gradient (can be omitted if preset colors are sufficient)
        gradientFill.SetTwoColorGradient(
            Color.Blue,               // first color
            Color.LightBlue,          // second color
            GradientStyleType.Horizontal,
            1);                       // variant

        // Set the gradient fill type and direction.
        // Here we use a linear gradient rotated 45 degrees, direction from upper left corner.
        gradientFill.SetGradient(
            GradientFillType.Linear,  // gradient type
            45.0,                     // angle (degrees) – applies to Linear type
            GradientDirectionType.FromUpperLeftCorner); // direction for Linear (used as angle)

        // Save the workbook to a file
        workbook.Save("WordArtGradientDirection.xlsx");
    }
}