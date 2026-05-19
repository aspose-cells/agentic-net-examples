using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePlacementDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target column index (0‑based)
                int targetColumnIndex = 3;

                // Add a rectangle shape; initial position will be adjusted later
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    0,                        // topRow
                    0,                        // top (pixel offset within the row)
                    targetColumnIndex,        // leftColumn
                    0,                        // left (pixel offset within the column)
                    100,                      // height in pixels
                    200);                     // width in pixels

                // Set the shape's column index and left offset (0 = column start)
                shape.UpperLeftColumn = targetColumnIndex;
                shape.Left = 0; // absolute position at the start of the column

                // Verify placement by comparing the shape's X property (offset from worksheet left border)
                if (shape.X == 0)
                {
                    Console.WriteLine("Shape placed correctly at X = " + shape.X);
                }
                else
                {
                    Console.WriteLine($"Shape placement mismatch. Expected X = 0, Actual X = {shape.X}");
                }

                // Save the workbook to verify the result
                string outputPath = "ShapePlacementDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}