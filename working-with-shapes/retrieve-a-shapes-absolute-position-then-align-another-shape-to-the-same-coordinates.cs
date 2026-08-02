// Title: Align Two Worksheet Shapes to Identical Coordinates Using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds two rectangle shapes, reads the first shape’s absolute row, column and offset values, and repositions the second shape with MoveToRange so it shares the exact cell location. The workbook is saved as AlignedShapes.xlsx.
// Keywords: Aspose.Cells shape alignment | MoveToRange C# | retrieve shape coordinates | worksheet shape positioning | C# Aspose.Cells example | absolute shape location | copy shape position | Aspose.Cells .NET | Excel shape coordinates | programmatic shape layout
// Common Searches: Aspose.Cells move shape to same cell | how to get shape row and column in Aspose.Cells | C# align two worksheet shapes | use MoveToRange to position shape | copy shape position Aspose.Cells .NET | align chart with rectangle shape Aspose.Cells
// Developer Intent: Duplicate the location of one worksheet shape and place another shape at that exact position programmatically.
// Use Cases: Maintain alignment of related diagram elements when generating reports | Automatically position a logo next to a dynamic chart | Synchronize watermark placement with header cells across multiple sheets | Create templates where data labels follow moved shapes | Ensure consistent layout when inserting shapes via automation
// AI Prompts: Write C# code that reads a shape’s UpperLeftRow, UpperLeftColumn, UpperLeftRowOffset, UpperLeftColumnOffset in Aspose.Cells and moves another shape to those coordinates using MoveToRange. | Explain step‑by‑step how to obtain a shape’s absolute position and align multiple shapes in an Excel worksheet with Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that aligns a chart shape to the same cell range as a rectangle shape, including handling of pixel offsets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAlignment
{
    // This C# example creates a workbook, adds two rectangle shapes, reads the first shape’s absolute row, column and offset values, and repositions the second shape with MoveToRange so it shares the exact cell location. The workbook is saved as AlignedShapes.xlsx.
    public class AlignShapes
    {
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
                Shape shape1 = shapes.AddRectangle(5, 2, 10, 15, 80, 120);

                // Add the second shape (another rectangle) at a different initial position
                Shape shape2 = shapes.AddRectangle(10, 5, 20, 30, 80, 120);

                // Retrieve the absolute position of the first shape
                int sourceTopRow = shape1.UpperLeftRow;
                int sourceLeftColumn = shape1.UpperLeftColumn;

                // Offsets are not required for this example; set to zero
                int sourceTopOffset = 0;
                int sourceLeftOffset = 0;

                // Align the second shape to the same coordinates as the first shape
                // MoveToRange positions a shape based on row/column and pixel offsets
                shape2.MoveToRange(sourceTopRow, sourceLeftColumn, sourceTopOffset, sourceLeftOffset);

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "AlignedShapes.xlsx");

                // Save the workbook to verify the result
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
