// Title: Copy a shape, rename it, and offset 10 points in Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to a worksheet, duplicate it with ShapeCollection.AddCopy, assign a new name, and shift the copy 10 pixels to the right before saving the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy shape | ShapeCollection.AddCopy | rename copied shape | offset shape position | C# Aspose.Cells shape example | .NET spreadsheet shape duplication | move shape 10 points | Aspose.Cells tutorial
// Common Searches: Aspose.Cells duplicate shape C# | How to rename a copied shape in Aspose.Cells | Shift copied shape horizontally Aspose.Cells | AddCopy method parameters example | Copy and move shape Aspose.Cells .NET
// Developer Intent: Create a duplicate of an existing worksheet shape, give the copy a distinct name, and move it 10 pixels to the right.
// Use Cases: Generate a series of offset shapes for a flow‑chart template. | Place a duplicated logo beside the original in a financial report. | Programmatically copy a chart placeholder, rename it, and align it with adjacent cells.
// AI Prompts: Show C# code that copies a shape in Aspose.Cells, renames the copy, and offsets it by 10 points horizontally. | Provide an Aspose.Cells .NET example using ShapeCollection.AddCopy to duplicate any shape and adjust its left position.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeCopyExample
{
    // Demonstrates how to add a rectangle shape to a worksheet, duplicate it with ShapeCollection.AddCopy, assign a new name, and shift the copy 10 pixels to the right before saving the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add an original rectangle shape
            // Parameters: upper left row, upper left column, top offset (pixels), left offset (pixels), width, height
            Shape originalShape = shapes.AddRectangle(2, 0, 2, 0, 130, 130);
            originalShape.Name = "OriginalShape";

            // Duplicate the original shape and offset it 10 points (pixels) to the right
            // Use the same row/column as the source shape, keep the same top offset,
            // increase the left offset by 10 pixels.
            Shape copiedShape = shapes.AddCopy(originalShape, 2, 2, 0, 10);
            copiedShape.Name = "CopiedShape";

            // Save the workbook to a file
            workbook.Save("ShapeCopyOffsetDemo.xlsx");
        }
    }
}
