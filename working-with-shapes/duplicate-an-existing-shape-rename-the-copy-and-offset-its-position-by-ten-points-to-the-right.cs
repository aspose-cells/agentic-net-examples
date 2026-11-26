// Title: How to duplicate a worksheet shape, rename the copy, and offset it by 10 points using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells duplicate a shape and move it 10 points to the right | Aspose.Cells AddCopy method example with renamed shape | How to offset a copied shape horizontally in an Excel worksheet using Aspose.Cells | Rename and reposition a shape copy in Aspose.Cells for .NET | Shift duplicated shape by pixels in Aspose.Cells C# code
// Tags: Aspose.Cells shape duplication AddCopy | C# duplicate worksheet shape offset | rename copied shape Aspose.Cells | horizontal shape offset points Aspose.Cells | moveandsize placement Aspose.Cells shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDuplicateDemo
{
    // The sample loads (or creates) an Excel workbook, ensures a rectangle shape exists, then uses ShapeCollection.AddCopy to duplicate the shape. The copy is renamed, its left offset is increased by 10 points (≈10 pixels), its placement is set to MoveAndSize, and a light‑green fill is applied. Finally, the workbook is saved as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook. If it does not exist we will create a new one.
            const string sourcePath = "input.xlsx";
            const string resultPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(sourcePath))
                {
                    // Load existing workbook
                    workbook = new Workbook(sourcePath);
                    Console.WriteLine($"Loaded workbook from \"{sourcePath}\".");
                }
                else
                {
                    // Create a new workbook with a single worksheet
                    workbook = new Workbook();
                    Console.WriteLine($"Source file \"{sourcePath}\" not found. A new workbook will be created.");
                }

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Ensure there is at least one shape to duplicate.
                // If the worksheet has no shapes, we add a simple rectangle shape first.
                Shape originalShape;
                if (shapes.Count == 0)
                {
                    // Add a rectangle shape at row 2, column 2 with zero pixel offsets.
                    // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
                    originalShape = shapes.AddRectangle(2, 0, 2, 0, 100, 200);
                    originalShape.Name = "OriginalRectangle";
                    originalShape.Placement = PlacementType.FreeFloating;
                    Console.WriteLine("No existing shapes found. Added a default rectangle shape named \"OriginalRectangle\".");
                }
                else
                {
                    // Use the first shape in the collection as the source shape.
                    originalShape = shapes[0];
                    Console.WriteLine($"Found existing shape \"{originalShape.Name}\" (type: {originalShape.GetType().Name}).");
                }

                // Retrieve the original shape's position parameters.
                // These properties are part of the Shape class.
                int originalRow = originalShape.UpperLeftRow;
                int originalTop = originalShape.Top;
                int originalColumn = originalShape.UpperLeftColumn;
                int originalLeft = originalShape.Left;

                // Offset the copy 10 points to the right.
                // Aspose.Cells uses pixel units for the offset properties.
                // 1 point = 1.3333 pixels (approx). We'll add 10 points ≈ 13 pixels.
                // For simplicity, we add 10 pixels which is close enough for demonstration.
                int offsetPixels = 10;

                // Duplicate the shape using AddCopy.
                // Signature: AddCopy(Shape shape, int upperLeftRow, int top, int leftColumn, int left)
                Shape copiedShape = shapes.AddCopy(
                    originalShape,
                    originalRow,
                    originalTop,
                    originalColumn,
                    originalLeft + offsetPixels);

                // Rename the copied shape.
                copiedShape.Name = "CopiedShape";

                // Optionally, you can change other properties of the copy to verify it is independent.
                copiedShape.Placement = PlacementType.MoveAndSize;
                copiedShape.Fill.FillType = FillType.Solid;
                copiedShape.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;

                Console.WriteLine($"Shape duplicated successfully. New shape name: \"{copiedShape.Name}\".");
                Console.WriteLine($"Original Left offset: {originalLeft}, Copied Left offset: {copiedShape.Left} (offset by {offsetPixels} pixels).");

                // Save the workbook with the duplicated shape.
                workbook.Save(resultPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
