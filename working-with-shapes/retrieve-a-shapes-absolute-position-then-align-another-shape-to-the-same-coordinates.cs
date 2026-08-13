// Title: Aspose.Cells .NET: Align One Shape to Another’s Position in a Worksheet
// Description: Demonstrates how to read a shape’s absolute row and column (UpperLeftRow, UpperLeftColumn) and move a second shape to the same cell using MoveToRange, without pixel offsets, then save the workbook.
// Keywords: Aspose.Cells align shapes | MoveToRange C# | shape absolute position | copy shape location | worksheet shape alignment
// Common Searches: how to align a shape to another shape in Aspose.Cells | retrieve shape row and column Aspose.Cells .NET | move shape to same cell as another shape | Aspose.Cells shape positioning example | copy shape coordinates Excel library
// Developer Intent: Place a target shape at the exact cell coordinates of a reference shape.
// Use Cases: Stack two rectangles so they occupy the same worksheet cell. | Position a picture to match the location of an existing chart. | Synchronize dynamically generated shapes based on a template shape.
// AI Prompts: Show C# code that gets a shape's UpperLeftRow and UpperLeftColumn and uses MoveToRange to align another shape in Aspose.Cells. | Provide an Aspose.Cells .NET example for aligning multiple shapes to a reference shape’s cell. | Explain how to move a shape to the same cell as another shape without using pixel offsets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to read a shape’s absolute row and column (UpperLeftRow, UpperLeftColumn) and move a second shape to the same cell using MoveToRange, without pixel offsets, then save the workbook.
    public class AlignShapeToAnother
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the first rectangle shape (source shape)
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape sourceShape = worksheet.Shapes.AddRectangle(2, 2, 10, 15, 100, 80);

                // Add a second rectangle shape (target shape) at a different initial position
                Shape targetShape = worksheet.Shapes.AddRectangle(5, 5, 20, 30, 100, 80);

                // Align the target shape to the same cell as the source shape.
                // Offsets are set to 0 because the Shape class does not expose pixel‑offset properties directly.
                targetShape.MoveToRange(sourceShape.UpperLeftRow, sourceShape.UpperLeftColumn, 0, 0);

                // Save the workbook to verify the result
                workbook.Save("AlignedShapes.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the example
    public class Program
    {
        public static void Main()
        {
            AlignShapeToAnother.Run();
        }
    }
}
