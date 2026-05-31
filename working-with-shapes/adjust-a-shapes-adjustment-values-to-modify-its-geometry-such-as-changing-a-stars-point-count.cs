using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class AdjustStarShapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a 10‑point star shape to the worksheet
                // Parameters: AutoShapeType, upper left row, upper left column, offset X, offset Y, width, height
                Shape star = worksheet.Shapes.AddAutoShape(AutoShapeType.Star10, 2, 2, 0, 0, 200, 200);

                // Access the geometry of the shape
                Geometry geometry = star.Geometry;

                // Adjust the first guide value if it exists; otherwise add a new guide
                if (geometry.ShapeAdjustValues.Count > 0)
                {
                    // For a star shape, the first adjustment typically controls the inner radius
                    geometry.ShapeAdjustValues[0].Value = 0.5; // Range usually 0.0‑1.0
                    Console.WriteLine("Adjusted first shape guide value to 0.5.");
                }
                else
                {
                    // Add a generic adjustment guide named "adj1"
                    int guideIndex = geometry.ShapeAdjustValues.Add("adj1", 0.5);
                    Console.WriteLine($"Added adjustment guide at index {guideIndex} with value 0.5.");
                }

                // Save the workbook
                string outputPath = "AdjustStarShapeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}