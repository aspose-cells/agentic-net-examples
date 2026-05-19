using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientTextboxDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Shape textBox = sheet.Shapes.AddTextBox(2, 0, 5, 5, 200, 80);
            textBox.Text = "Gradient Filled TextBox";

            // Set the fill type of the textbox to Gradient to access GradientFill properties
            textBox.Fill.FillType = FillType.Gradient;

            // Retrieve the GradientFill object
            GradientFill gradientFill = textBox.Fill.GradientFill;

            // Define two custom colors for the gradient
            Color startColor = Color.FromArgb(255, 255, 200, 0);   // Orange
            Color endColor   = Color.FromArgb(255, 0, 120, 215);   // Blue

            // Apply a two‑color gradient (horizontal style, variant 1)
            gradientFill.SetTwoColorGradient(startColor, endColor, GradientStyleType.Horizontal, 1);

            // Set the angle of the linear gradient (e.g., 45 degrees)
            gradientFill.Angle = 45f;

            // Save the workbook to a file
            workbook.Save("GradientTextboxDemo.xlsx");
        }
    }
}