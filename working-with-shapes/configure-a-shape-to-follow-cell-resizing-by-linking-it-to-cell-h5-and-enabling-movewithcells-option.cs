// Title: Aspose.Cells for .NET – Link a Shape to Cell H5 and Enable Move‑and‑Size with Cells
// Description: Shows how to insert a rectangle shape, bind it to cell H5 via the LinkedCell property, set its Placement to MoveAndSize so it follows cell resizing, and save the workbook.
// Keywords: Aspose.Cells | .NET | shape LinkedCell | PlacementType.MoveAndSize | move and size with cells | rectangle shape Excel | cell H5 binding | programmatic shape placement | Excel shape resizing
// Common Searches: Aspose.Cells link shape to cell | Set shape placement MoveAndSize in C# | Bind a shape to a specific cell using Aspose.Cells | Move and resize shape with cell Aspose.Cells .NET | C# example for shape LinkedCell property
// Developer Intent: Bind a worksheet shape to cell H5 and have it move and resize together with that cell.
// Use Cases: Create a dynamic dashboard where icons stay aligned with key metrics as rows/columns are adjusted. | Generate reports with status symbols attached to cells that automatically adapt to layout changes. | Design templates with placeholder shapes tied to fixed cells, preserving design integrity during user edits.
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle, links it to H5, and sets Placement to MoveAndSize. | Explain how to modify an existing shape’s LinkedCell and Placement at runtime in an Aspose.Cells workbook. | Show how to read a shape’s LinkedCell value and programmatically adjust its size based on cell dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a rectangle shape, bind it to cell H5 via the LinkedCell property, set its Placement to MoveAndSize so it follows cell resizing, and save the workbook.
class ShapeCellLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

            // Link the shape to cell H5 so it follows that cell
            shape.LinkedCell = "H5";

            // Enable "Move and size with cells" behavior
            shape.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("ShapeLinkedToH5.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
