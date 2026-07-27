// Title: C# – Add a rectangle shape with custom pixel dimensions and anchor it to a cell range using Aspose.Cells
// Description: Creates a new workbook, adds a rectangle shape sized 200 px × 100 px, anchors it to the B2:D5 range with MoveToRange, and saves the file as AnchoredShape.xlsx.
// Keywords: Aspose.Cells AddShape | C# shape width height pixels | MoveToRange anchor shape | Aspose.Cells rectangle shape example | Excel shape positioning .NET | custom sized shape Aspose.Cells
// Common Searches: Aspose.Cells add rectangle shape with pixel size | How to anchor a shape to a cell range in C# | MoveToRange method usage Aspose.Cells | Set shape width and height in pixels Aspose.Cells | C# code to place a shape over B2:D5
// Developer Intent: Insert a 200 px × 100 px rectangle and bind it to the B2:D5 range in a new Excel workbook.
// Use Cases: Add a banner that stays aligned with a header block. | Place a logo with exact pixel dimensions over a designated cell area. | Overlay a colored shape on a table so it moves and resizes with the data range.
// AI Prompts: Generate C# code to add an ellipse of 150 px width and 80 px height anchored to cells C3:E6 with Aspose.Cells. | Show how to modify the fill color and line style of a rectangle after anchoring it using MoveToRange. | Create a reusable method that accepts width, height, and a cell range, then adds a shape with those parameters to a worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    // Creates a new workbook, adds a rectangle shape sized 200 px × 100 px, anchors it to the B2:D5 range with MoveToRange, and saves the file as AnchoredShape.xlsx.
    public class AddAnchoredShape
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add a rectangle shape with custom width (200 pixels) and height (100 pixels)
                // Initial position offsets are set to 0; they will be adjusted by MoveToRange
                Shape shape = shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    0,   // topRow (temporary)
                    0,   // top offset in pixels
                    0,   // leftColumn (temporary)
                    0,   // left offset in pixels
                    100, // height in pixels
                    200  // width in pixels
                );

                // Anchor the shape to a specific cell range, e.g., B2:D5
                // Rows and columns are zero‑based indexes: B2 -> row 1, column 1; D5 -> row 4, column 3
                shape.MoveToRange(1, 1, 4, 3);

                // Define output file path
                string outputPath = "AnchoredShape.xlsx";

                // Save the workbook with the added shape
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            AddAnchoredShape.Run();
        }
    }
}
