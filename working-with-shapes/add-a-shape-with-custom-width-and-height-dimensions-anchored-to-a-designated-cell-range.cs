// Title: Create a rectangle shape with specific width and height anchored to a cell range in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape with custom size to a specific cell range | how to anchor a shape to cells B2:C3 and set its width and height using Aspose.Cells | set PlacementType.MoveAndSize for a shape added with AddShape in Aspose.Cells .NET | define shape dimensions in points when adding a rectangle to an Excel worksheet with Aspose.Cells | C# example for adding a shape that moves and resizes with cells in Aspose.Cells
// Tags: AddShape rectangle custom dimensions Aspose.Cells | anchor shape to cell range C# | PlacementType.MoveAndSize Aspose.Cells | shape width height points Aspose.Cells | Excel shape positioning B2:C3 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a rectangle shape of 150 × 80 points anchored to the B2:C3 range, configuring it to move and size with the cells, and saving the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range to which the shape will be anchored (e.g., B2:C3)
            // Row and column indexes are zero‑based.
            int upperRow = 1;      // Row for cell B2
            int upperColumn = 1;   // Column for cell B2
            int lowerRow = 2;      // Row for cell C3
            int lowerColumn = 2;   // Column for cell C3

            // Add a rectangle shape anchored to the specified cell range
            // Parameters: drawing type, upper row, upper column, lower row, lower column, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperRow, upperColumn,
                lowerRow, lowerColumn,
                150, // width in points
                80   // height in points
            );

            // Ensure the shape moves and sizes with the cells
            shape.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
