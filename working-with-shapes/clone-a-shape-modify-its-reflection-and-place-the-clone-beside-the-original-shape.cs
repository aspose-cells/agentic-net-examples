// Title: Clone a rectangle shape, change its reflection effect, and position the clone beside the original using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clone rectangle shape and apply a different reflection effect | how to place a copied shape next to the original in an Excel worksheet using Aspose.Cells | C# code to duplicate a shape with custom reflection settings in Aspose.Cells | move cloned shape columns to the right with Aspose.Cells API
// Tags: clone rectangle shape Aspose.Cells | reflection effect properties Aspose.Cells shape | add rectangle shape to worksheet C# | position cloned shape beside original Aspose.Cells | modify shape reflection Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCloneReflectionDemo
{
    // The example creates a workbook, adds a rectangle shape with a half‑reflection effect, clones the shape five columns to the right, applies a different full‑reflection effect to the clone, and saves the file as ShapeCloneReflectionDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Add an original rectangle shape
                // -----------------------------------------------------------------
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape originalShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 120, 80);
                originalShape.Name = "OriginalRectangle";

                // Set a reflection effect on the original shape
                ReflectionEffect originalReflection = originalShape.Reflection;
                originalReflection.Type = ReflectionEffectType.HalfReflectionTouching;
                originalReflection.Size = 60;          // 60% of the shape height
                originalReflection.Blur = 5;           // moderate blur
                originalReflection.Transparency = 0.3; // 30% transparent
                originalReflection.Distance = 10;      // distance from the shape
                originalReflection.Direction = 90;    // vertical reflection
                originalReflection.FadeDirection = 90;
                originalReflection.RotWithShape = true;

                // -----------------------------------------------------------------
                // 2. Clone the shape and place the clone beside the original
                // -----------------------------------------------------------------
                // Determine the original shape's position and size
                int originalRow = originalShape.UpperLeftRow;
                int originalColumn = originalShape.UpperLeftColumn;
                // Offsets are not required for this demo; use 0
                int originalRowOffset = 0;
                int originalColumnOffset = 0;
                int shapeWidth = (int)originalShape.Width;
                int shapeHeight = (int)originalShape.Height;

                // Move the clone 5 columns to the right
                int cloneColumnOffset = 5;
                Shape clonedShape = worksheet.Shapes.AddRectangle(
                    originalRow,
                    originalColumn + cloneColumnOffset,
                    originalRowOffset,
                    originalColumnOffset,
                    shapeWidth,
                    shapeHeight);
                clonedShape.Name = "ClonedRectangle";

                // -----------------------------------------------------------------
                // 3. Modify the reflection of the cloned shape
                // -----------------------------------------------------------------
                ReflectionEffect clonedReflection = clonedShape.Reflection;
                clonedReflection.Type = ReflectionEffectType.FullReflectionTouching;
                clonedReflection.Size = 80;           // larger reflection
                clonedReflection.Blur = 2;            // less blur for a sharper look
                clonedReflection.Transparency = 0.1;  // almost opaque
                clonedReflection.Distance = 5;        // closer to the shape
                clonedReflection.Direction = 45;      // diagonal reflection
                clonedReflection.FadeDirection = 45;
                clonedReflection.RotWithShape = false;

                // -----------------------------------------------------------------
                // 4. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "ShapeCloneReflectionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Original shape and cloned shape with distinct reflection effects have been created.");
            }
            catch (Exception ex)
            {
                // Any exception will be printed to the console – do not hide failures.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
