using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientShapeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 150, 100);

            // Set the fill type of the shape to Gradient to enable gradient properties
            shape.Fill.FillType = FillType.Gradient;

            // Obtain the GradientFill object from the shape's Fill
            GradientFill gradientFill = shape.Fill.GradientFill;

            // Define a linear gradient with two colors (light sky blue to dark blue)
            // Use Horizontal style and the first variant
            gradientFill.SetTwoColorGradient(
                Color.LightSkyBlue,   // First color
                Color.DarkBlue,       // Second color
                GradientStyleType.Horizontal,
                1);

            // Set the gradient type to Linear and specify the angle (e.g., 45 degrees)
            gradientFill.SetGradient(GradientFillType.Linear, 45.0, GradientDirectionType.FromCenter);

            // Save the workbook with the gradient-filled shape
            workbook.Save("LinearGradientShape.xlsx");
        }
    }
}