using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class AlignShapeToAnotherShape
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

                // Get the shapes collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add the first shape (a rectangle) at a specific position
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                int sourceRow = 2;
                int sourceColumn = 2;
                int sourceRowOffset = 10;   // top offset (in pixels)
                int sourceColumnOffset = 15; // left offset (in pixels)
                Shape sourceShape = shapes.AddRectangle(sourceRow, sourceColumn, sourceRowOffset, sourceColumnOffset, 100, 200);

                // Add a second shape (another rectangle) at a different initial position
                Shape targetShape = shapes.AddRectangle(5, 5, 30, 40, 100, 200);

                // Align the target shape to the same position as the source shape
                targetShape.MoveToRange(sourceRow, sourceColumn, sourceRowOffset, sourceColumnOffset);

                // Verify that the positions match by printing them
                Console.WriteLine($"Source Shape - Row:{sourceShape.UpperLeftRow}, RowOffset:{sourceRowOffset}, Column:{sourceShape.UpperLeftColumn}, ColumnOffset:{sourceColumnOffset}");
                Console.WriteLine($"Target Shape - Row:{targetShape.UpperLeftRow}, RowOffset:{sourceRowOffset}, Column:{targetShape.UpperLeftColumn}, ColumnOffset:{sourceColumnOffset}");

                // Save the workbook to a file
                try
                {
                    workbook.Save("AlignedShapes.xlsx");
                    Console.WriteLine("Workbook saved as AlignedShapes.xlsx");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}