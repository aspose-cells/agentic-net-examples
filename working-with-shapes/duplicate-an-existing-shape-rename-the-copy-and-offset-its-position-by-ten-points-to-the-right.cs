using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace DuplicateShapeDemo
{
    class DuplicateShapeExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the shapes collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add an original rectangle shape
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Shape original = shapes.AddRectangle(2, 0, 2, 0, 130, 130);

                // Duplicate the shape, shifting it 10 pixels to the right
                // Offsets are specified in pixels; use 0 for row offset and 10 for column offset
                Shape copy = shapes.AddCopy(
                    original,
                    original.UpperLeftRow,   // same top row
                    0,                        // same vertical offset
                    original.UpperLeftColumn,// same left column
                    10);                      // increase horizontal offset by 10 pixels

                // Rename the copied shape
                copy.Name = "RectangleCopy";

                // Define output file path
                string outputPath = "DuplicateShape.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}