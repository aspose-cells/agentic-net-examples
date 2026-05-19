using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
                Shape shape2 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

                // Record initial Z-order positions
                Console.WriteLine($"Initial ZOrderPosition of shape1: {shape1.ZOrderPosition}");
                Console.WriteLine($"Initial ZOrderPosition of shape2: {shape2.ZOrderPosition}");

                // Move shape2 to the front if not already at front
                if (shape2.ZOrderPosition < worksheet.Shapes.Count - 1)
                {
                    shape2.ToFrontOrBack(1);
                }
                Console.WriteLine($"After moving shape2 to front, ZOrderPosition: {shape2.ZOrderPosition}");

                // Move shape1 to the back if not already at back
                if (shape1.ZOrderPosition > 0)
                {
                    shape1.ToFrontOrBack(-1);
                }
                Console.WriteLine($"After moving shape1 to back, ZOrderPosition: {shape1.ZOrderPosition}");

                // Save the workbook
                string outputPath = "ShapeZOrderDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}