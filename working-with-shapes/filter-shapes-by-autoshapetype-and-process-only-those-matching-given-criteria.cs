using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeFilterDemo
{
    class ShapeFilterExample
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add shapes of various AutoShapeTypes
                sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 2, 2, 100, 50, 0, 0);
                sheet.Shapes.AddAutoShape(AutoShapeType.Oval, 5, 2, 100, 50, 0, 0);
                sheet.Shapes.AddAutoShape(AutoShapeType.Star5, 8, 2, 100, 50, 0, 0);

                // Define the AutoShapeType to filter (e.g., Rectangle)
                AutoShapeType targetType = AutoShapeType.Rectangle;

                // Iterate through all shapes and process those matching the criteria
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.AutoShapeType == targetType)
                    {
                        // Example processing: change fill color and rotation
                        shape.Fill.SolidFill.Color = Color.LightGreen;
                        shape.RotationAngle = 15;
                    }
                }

                // Save the workbook
                string outputPath = "FilteredShapes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing: {ex.Message}");
            }
        }
    }
}