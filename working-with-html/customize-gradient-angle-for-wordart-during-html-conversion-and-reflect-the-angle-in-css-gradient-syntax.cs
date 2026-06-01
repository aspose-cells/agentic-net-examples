using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtGradientHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape with a preset style (style 7 has a gradient fill)
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Gradient WordArt",
                2,   // top row
                10,  // top offset (pixels)
                2,   // left column
                10,  // left offset (pixels)
                100, // height (pixels)
                400  // width (pixels)
            );

            // Ensure the shape's fill is set to gradient to access GradientFill properties
            wordArt.Fill.FillType = FillType.Gradient;
            GradientFill gradientFill = wordArt.Fill.GradientFill;

            // Set the gradient type to linear and define the angle (e.g., 45 degrees)
            double angle = 45.0;
            gradientFill.SetGradient(GradientFillType.Linear, angle, GradientDirectionType.FromCenter);
            // Explicitly set the Angle property as well (both achieve the same result)
            gradientFill.Angle = (float)angle;

            // Define the two colors for the gradient
            gradientFill.SetTwoColorGradient(
                Color.Blue,          // start color
                Color.LightBlue,     // end color
                GradientStyleType.Horizontal,
                1                    // variant
            );

            // Convert the workbook to HTML. The generated CSS will contain the gradient with the specified angle.
            workbook.Save("WordArtGradient.html", SaveFormat.Html);

            Console.WriteLine("Workbook saved as HTML with WordArt gradient angle reflected in CSS.");
        }
    }
}