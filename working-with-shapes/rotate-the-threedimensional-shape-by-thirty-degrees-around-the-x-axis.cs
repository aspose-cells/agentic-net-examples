using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

namespace AsposeCellsExamples
{
    public class RotateShapeAroundXAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 2, 200, 100, 100);

                // Access the 3‑D format of the shape
                ThreeDFormat threeDFormat = shape.ThreeDFormat;

                // Set rotation around the X‑axis to 30 degrees
                threeDFormat.RotationX = 30;

                // Additional 3‑D properties to make the rotation visible
                threeDFormat.ExtrusionHeight = 20;          // give the shape depth
                threeDFormat.ContourWidth = 2;              // outline thickness
                threeDFormat.PresetCameraType = PresetCameraType.OrthographicFront; // view angle

                string outputPath = "ShapeRotatedAroundX30Degrees.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Shape rotated 30 degrees around X‑axis and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RotateShapeAroundXAxis.Run();
        }
    }
}