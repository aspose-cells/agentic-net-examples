// Title: Place a rectangle shape at a specific column using pixel offset calculated from column width in Aspose.Cells for .NET
// Common Searches: how to calculate pixel offset for a shape based on column width in Aspose.Cells C# | Aspose.Cells set shape left position using column index and GetColumnWidthPixel | add rectangle shape to worksheet and position it at column C in .NET | C# Aspose.Cells place shape at absolute pixel coordinates | determine column width in pixels for shape placement using Aspose.Cells
// Tags: shape left offset calculation Aspose.Cells | insert rectangle shape worksheet C# | retrieve column width in pixels Aspose.Cells | absolute shape positioning using column index | set shape size and location Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, obtains the pixel width of a target column, computes an absolute X offset by multiplying the column index with the column width, adds a rectangle shape, sets its dimensions, positions it using the calculated pixel offset, writes verification data to a cell, and saves the file as ShapePlacement.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Column index (0‑based) where we want the shape to start (C = index 2)
            int targetColumn = 2;

            // Retrieve the width of the target column in pixels
            int columnWidthPixels = sheet.Cells.GetColumnWidthPixel(targetColumn);

            // Calculate the absolute X offset: column index * column width
            int absoluteXOffset = targetColumn * columnWidthPixels;

            // Add a rectangle shape; initial size is zero (will be set later)
            // AddShape returns the created Shape object in recent API versions
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,   // upper left row
                0,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                0,   // height (points)
                0);  // width (points)

            // Set shape size (width and height in points)
            shape.Width = 100;   // 100 points wide
            shape.Height = 50;   // 50 points high

            // Position the shape using absolute pixel offsets
            shape.Top = 0;                     // top offset in pixels
            shape.Left = absoluteXOffset;      // left offset in pixels

            // Write verification data to a cell
            sheet.Cells["A1"].PutValue(
                $"Shape placed at column index {targetColumn} with X offset {absoluteXOffset} pixels.");

            // Save the workbook
            workbook.Save("ShapePlacement.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
