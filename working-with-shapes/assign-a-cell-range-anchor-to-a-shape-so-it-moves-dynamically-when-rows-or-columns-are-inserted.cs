// Title: Aspose.Cells .NET – Anchor a Shape to a Cell Range with Two‑Cell Anchor (Move & Resize)
// Description: Demonstrates how to add a rectangle shape, set its AnchorType to TwoCellAnchor, bind it to the B2:D5 range using MoveToRange (zero‑based indices), and configure Placement to MoveAndSize so the shape follows inserted rows or columns.
// Keywords: Aspose.Cells shape anchor | TwoCellAnchor .NET | shape MoveAndSize | bind shape to cell range | dynamic shape positioning | C# Aspose.Cells example
// Common Searches: Aspose.Cells bind shape to range | TwoCellAnchor shape Aspose.Cells C# | shape moves when rows inserted Aspose | how to anchor a rectangle to cells in Aspose.Cells | PlacementType.MoveAndSize example
// Developer Intent: Attach a shape to a specific cell block so it automatically moves and resizes when the worksheet layout changes.
// Use Cases: Keep a rectangle aligned with B2:D5 after adding rows or columns. | Anchor charts, images, or text boxes to a dynamic range for reporting templates. | Create printable forms where shapes must expand or contract with cell data.
// AI Prompts: Generate C# code that anchors a shape to B2:D5 using Aspose.Cells and ensures it moves with inserted rows. | Explain the impact of PlacementType.MoveAndSize on a TwoCellAnchor‑bound shape. | Show how to use MoveToRange with zero‑based indices to bind a shape to a cell range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape, set its AnchorType to TwoCellAnchor, bind it to the B2:D5 range using MoveToRange (zero‑based indices), and configure Placement to MoveAndSize so the shape follows inserted rows or columns.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

        // Use a two‑cell anchor so the shape is bound to a cell range
        shape.AnchorType = ShapeAnchorType.TwoCellAnchor;

        // Anchor the shape to the range B2:D5 (rows and columns are zero‑based)
        shape.MoveToRange(1, 1, 4, 3);

        // Make the shape move and resize together with the cells
        shape.Placement = PlacementType.MoveAndSize;

        // Save the workbook
        workbook.Save("ShapeWithRangeAnchor.xlsx");
    }
}
