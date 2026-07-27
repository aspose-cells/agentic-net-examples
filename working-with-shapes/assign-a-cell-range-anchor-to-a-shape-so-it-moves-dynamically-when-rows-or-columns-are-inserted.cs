// Title: Anchor a Shape to a Cell Range in Aspose.Cells for .NET (C#) – Move and Resize with Inserted Rows/Columns
// Description: Demonstrates how to add a rectangle shape, anchor it to the range B2:D5 using MoveToRange, set PlacementType.MoveAndSize, and save the workbook so the shape automatically follows any row or column insertions.
// Keywords: Aspose.Cells shape anchoring | MoveToRange C# | PlacementType.MoveAndSize | shape moves with rows | shape resizes with columns | .NET Excel shape range | dynamic shape positioning
// Common Searches: Aspose.Cells anchor shape to range C# | MoveToRange example Aspose.Cells | PlacementType.MoveAndSize behavior | shape follows inserted rows Aspose.Cells | how to bind shape to cell range .NET
// Developer Intent: Bind a worksheet shape to a specific cell range so it automatically moves and resizes when the sheet layout changes.
// Use Cases: Keep a highlight box aligned with a data table as new rows are added. | Attach a company logo to header cells that shift when columns are inserted. | Anchor a comment or note box to a merged range that expands with the range.
// AI Prompts: Generate C# code that anchors a picture to range A1:C3 with PlacementType.MoveAndSize using Aspose.Cells. | Show how to change the anchor of an existing shape to a new range after loading a workbook. | Explain the differences between PlacementType.MoveAndSize, Move, and FreeFloating when rows or columns are inserted.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape, anchor it to the range B2:D5 using MoveToRange, set PlacementType.MoveAndSize, and save the workbook so the shape automatically follows any row or column insertions.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, width, height, upper left offset X, upper left offset Y
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

        // Anchor the shape to a cell range (B2:D5)
        // MoveToRange uses zero‑based row/column indices:
        //   topRow = 1 (row 2), leftColumn = 1 (column B)
        //   bottomRow = 4 (row 5), rightColumn = 3 (column D)
        shape.MoveToRange(1, 1, 4, 3);

        // Ensure the shape moves and resizes with the cells it is anchored to
        shape.Placement = PlacementType.MoveAndSize;

        // Save the workbook
        workbook.Save("ShapeAnchorDemo.xlsx");
    }
}
