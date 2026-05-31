using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class MoveShapeToFrontDemo
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping shapes to demonstrate z‑order
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

            // Bring shape2 to the front by moving it forward 1 position (positive integer)
            shape2.ToFrontOrBack(1);

            // Optionally, send shape1 to the back for contrast
            shape1.ToFrontOrBack(-1);

            // Define output file path
            string outputPath = "MoveShapeToFrontDemo.xlsx";

            // Save the workbook (overwrite if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
    }
}