// Title: Link a rectangle shape to a SEQUENCE dynamic array cell and enable MoveAndSize behavior with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind rectangle shape text to a cell containing a SEQUENCE dynamic array | Set shape placement to MoveAndSize for a shape linked to a dynamic array cell in Aspose.Cells | How to make a shape follow and resize with a dynamic array range in an Excel workbook using Aspose.Cells | C# update shape position automatically when SEQUENCE array expands with Aspose.Cells | Link shape to a cell with a dynamic array formula and keep it synchronized in .NET
// Tags: link shape to dynamic array cell Aspose.Cells | shape placement MoveAndSize Aspose.Cells | rectangle shape text formula binding C# | SEQUENCE function shape linking .NET | shape resize with cell range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a SEQUENCE dynamic array formula into cell A1, adds a rectangle shape at a specific location, sets the shape's Placement to MoveAndSize so it moves and resizes with the cells it occupies, binds the shape's displayed text to cell A1 using a formula, assigns alternative text for identification, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a dynamic array formula in cell A1 (example: SEQUENCE)
            sheet.Cells["A1"].Formula = "=SEQUENCE(5,1,1,1)";

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,          // upper left row
                5,          // upper left column
                0,          // row offset (pixels)
                0,          // column offset (pixels)
                100,        // height (pixels)
                200);       // width (pixels)

            // Make the shape move and resize with the cells it occupies
            shape.Placement = PlacementType.MoveAndSize;

            // Link the shape's displayed text to the first cell of the array (A1)
            shape.Text = "=A1";

            // Optional: give the shape an alternative text for identification
            shape.AlternativeText = "DynamicArrayLinkedShape";

            // Save the workbook
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
