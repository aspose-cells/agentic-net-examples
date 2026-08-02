// Title: Aspose.Cells for .NET – Link a Rectangle Shape to cell H5 and enable Move‑And‑Size
// Description: C# example that creates a workbook, adds a rectangle shape, links it to cell H5, sets its placement to MoveAndSize so the shape automatically resizes with the cell, and saves the file as ShapeLinkedToH5.xlsx.
// Keywords: Aspose.Cells C# shape linked cell | MoveAndSize placement Aspose.Cells | rectangle shape cell H5 | shape follow cell resizing .NET | Aspose.Cells shape placement
// Common Searches: Aspose.Cells link shape to a specific cell | How to make a shape resize with a cell in Aspose.Cells | C# set shape placement MoveAndSize Aspose | Rectangle shape follow cell H5 Aspose.Cells example
// Developer Intent: Attach a rectangle shape to cell H5 and have it automatically move and resize when the cell dimensions change.
// Use Cases: Dynamic dashboards where highlight boxes stay aligned with key data cells. | Automated report templates that preserve shape positions after column or row adjustments. | Interactive spreadsheets with comment or annotation shapes that remain bound to their target cells.
// AI Prompts: Generate C# code that links a shape to cell B2 and applies MoveAndSize placement using Aspose.Cells. | Explain the effect of PlacementType.MoveAndSize on shape behavior during row or column resizing. | Show how to update the LinkedCell property of an existing Aspose.Cells shape to a new address.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a rectangle shape, links it to cell H5, sets its placement to MoveAndSize so the shape automatically resizes with the cell, and saves the file as ShapeLinkedToH5.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: upper left row, upper left column, row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 50, 100);

            // Link the shape to cell H5 so it follows that cell
            shape.LinkedCell = "H5";

            // Enable move‑and‑size with cells (the shape will resize when the cell changes size)
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
