// Title: Aspose.Cells for .NET: Add a Rectangle Shape and Position It Using Column Pixel Width (C#)
// Description: Demonstrates how to create a workbook, retrieve a column's pixel width, compute an absolute X offset (column index × width), add a rectangle shape, set its X coordinate, apply MoveAndSize placement, output diagnostic values, and save the file as an Excel workbook.
// Keywords: Aspose.Cells C# shape positioning | absolute X offset column pixel width | add rectangle shape Aspose.Cells | PlacementType.MoveAndSize | column width pixels Aspose.Cells | Excel shape alignment .NET | shape X property Aspose.Cells
// Common Searches: Aspose.Cells set shape X coordinate by column width | C# calculate column pixel offset for shape placement | How to anchor a shape to a specific column in Aspose.Cells | MoveAndSize placement for shapes in Aspose.Cells .NET | Retrieve column width in pixels Aspose.Cells example
// Developer Intent: Place a rectangle shape at a precise X coordinate derived from a column's pixel width.
// Use Cases: Align a banner shape with the start of column D so it stays aligned after column resizing. | Create a chart placeholder that moves and resizes together with cells in column B. | Build a template where shapes are locked to specific columns, preserving layout during user edits.
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape and position it at column 5 based on the column's pixel width, with MoveAndSize placement. | Explain step‑by‑step how to obtain a column's width in pixels and convert it to an absolute X offset for shape positioning in Aspose.Cells. | Provide a reusable method that updates the X positions of existing shapes when column widths change in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, retrieve a column's pixel width, compute an absolute X offset (column index × width), add a rectangle shape, set its X coordinate, apply MoveAndSize placement, output diagnostic values, and save the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Column index to base the absolute position on (0‑based, e.g., column C = 2)
        int targetColumn = 3;

        // Retrieve the width of the target column in pixels
        int columnWidthPx = sheet.Cells.GetColumnWidthPixel(targetColumn);

        // Calculate the absolute X offset (column index * column width)
        int absoluteX = targetColumn * columnWidthPx;

        // Add a rectangle shape at the top‑left corner of the sheet
        RectangleShape shape = sheet.Shapes.AddRectangle(
            topRow: 0,    // upper left row index
            top: 0,       // vertical offset in pixels
            leftColumn: 0,// upper left column index
            left: 0,      // horizontal offset in pixels
            height: 100,  // height in pixels
            width: 200);  // width in pixels

        // Set the shape's absolute X position using the computed offset
        shape.X = absoluteX;

        // Optionally set placement to move and size with cells
        shape.Placement = PlacementType.MoveAndSize;

        // Test and display placement information
        Console.WriteLine($"Target column index: {targetColumn}");
        Console.WriteLine($"Column width (pixels): {columnWidthPx}");
        Console.WriteLine($"Computed absolute X offset: {absoluteX}");
        Console.WriteLine($"Shape X property after setting: {shape.X}");

        // Save the workbook
        workbook.Save("ShapeAbsolutePositionDemo.xlsx");
    }
}
