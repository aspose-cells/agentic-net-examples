using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeToFrontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
                Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

                // Bring shape2 to the front of the Z‑order (positive value)
                shape2.ToFrontOrBack(1);

                // Optionally, send shape1 to the back of the Z‑order (negative value)
                shape1.ToFrontOrBack(-1);

                // Save the workbook with the updated shape ordering
                string outputPath = "ShapeToFrontDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeToFrontDemo.Run();
        }
    }
}