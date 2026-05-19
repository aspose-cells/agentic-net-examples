using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class BringShapeToFront
    {
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
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                Shape shapeBack = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
                Shape shapeFront = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

                // Bring the second shape to the front of the Z‑order stack
                // 1 = bring to front, 0 = send to back
                shapeFront.ToFrontOrBack(1);

                // Define output path and ensure directory exists
                string outputPath = "ShapeFrontDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors that may occur during processing
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}