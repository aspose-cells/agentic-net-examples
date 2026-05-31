using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class CalloutTipAngleAdjustment
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a callout shape (Right Arrow Callout) to the worksheet
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape callout = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);

                // Access the geometry of the shape which holds adjustment guides
                Geometry geometry = callout.Geometry;

                // The tip angle of a callout shape is typically the third adjustment guide (index 2)
                int tipAngleIndex = 2; // zero‑based index

                // Ensure the shape has enough adjustment guides
                if (geometry.ShapeAdjustValues.Count > tipAngleIndex)
                {
                    // Set the tip angle to a desired value (e.g., 0.25 = 25% of the possible range)
                    geometry.ShapeAdjustValues[tipAngleIndex].Value = 0.25;
                    Console.WriteLine($"Tip angle adjustment (index {tipAngleIndex}) set to 0.25");
                }
                else
                {
                    Console.WriteLine("The shape does not contain the expected adjustment guide for tip angle.");
                }

                // Define output file path
                string outputPath = "CalloutTipAngleAdjustment.xlsx";

                // Save the workbook with the modified shape
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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
            CalloutTipAngleAdjustment.Run();
        }
    }
}