using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCloneReflectionDemo
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

                // Add an original rectangle shape
                // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, width, height
                Shape originalShape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 80);

                // Set reflection properties on the original shape for visual reference
                ReflectionEffect originalReflection = originalShape.Reflection;
                originalReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
                originalReflection.Transparency = 0.3;
                originalReflection.Size = 60;
                originalReflection.Blur = 2;
                originalReflection.Distance = 5;

                // Clone the original shape and place the copy beside it
                // We'll place the copy two columns to the right of the original shape
                int originalLeftColumn = originalShape.UpperLeftColumn;
                int originalTopRow = originalShape.UpperLeftRow;

                // AddCopy returns the new shape instance
                Shape clonedShape = worksheet.Shapes.AddCopy(
                    originalShape,
                    originalTopRow,          // same top row
                    0,                       // same vertical offset
                    originalLeftColumn + 5, // shift right by 5 columns
                    0);                      // same horizontal offset

                // Modify the reflection of the cloned shape
                ReflectionEffect clonedReflection = clonedShape.Reflection;
                clonedReflection.Type = ReflectionEffectType.HalfReflectionTouching;
                clonedReflection.Transparency = 0.6;
                clonedReflection.Size = 40;
                clonedReflection.Blur = 4;
                clonedReflection.Distance = 8;

                // Optionally, flip the cloned shape horizontally to emphasize the difference
                clonedShape.IsFlippedHorizontally = true;

                // Save the workbook to a file
                workbook.Save("ShapeCloneReflectionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}