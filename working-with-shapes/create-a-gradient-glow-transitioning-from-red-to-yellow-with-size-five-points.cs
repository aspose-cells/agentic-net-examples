using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class GradientGlowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will receive the gradient fill and glow effect
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

                // Set the fill type to gradient
                shape.Fill.FillType = FillType.Gradient;

                // Configure two‑color gradient: red to yellow, horizontal style
                GradientFill gradientFill = shape.Fill.GradientFill;
                gradientFill.SetTwoColorGradient(
                    Color.Red,
                    Color.Yellow,
                    GradientStyleType.Horizontal,
                    1);

                // Apply a glow effect: size 5 points, color yellow
                shape.Glow.Size = 5;
                CellsColor glowColor = workbook.CreateCellsColor();
                glowColor.Color = Color.Yellow;
                shape.Glow.Color = glowColor;

                // Save the workbook
                string outputPath = "GradientGlowDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            GradientGlowDemo.Run();
        }
    }
}