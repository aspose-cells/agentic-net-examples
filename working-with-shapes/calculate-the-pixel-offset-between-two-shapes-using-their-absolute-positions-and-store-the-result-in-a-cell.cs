// Title: C# – Compute pixel offset between two shapes and write results to cells using Aspose.Cells
// Description: Creates a workbook, adds two rectangle shapes with absolute pixel positions, calculates the horizontal (X) and vertical (Y) pixel differences, optionally derives the Euclidean distance, writes these values to cells A1‑B3, and saves the file as ShapeOffsetResult.xlsx.
// Keywords: Aspose.Cells C# shape offset | pixel distance between shapes | Left Top properties Aspose.Cells | calculate shape separation | store values in Excel cells | AddRectangle Aspose.Cells | shape alignment check | Excel workbook automation
// Common Searches: Aspose.Cells get shape left top pixel values | calculate offset between two shapes in .NET | write shape distance to Excel cell using Aspose.Cells | C# pixel distance between rectangle shapes | store shape coordinates in worksheet cells
// Developer Intent: Find the X/Y pixel difference (and optional Euclidean distance) between two shapes and record the numbers in worksheet cells.
// Use Cases: Validate diagram layout by comparing relative positions of two objects and logging the offsets. | Generate a report that lists X/Y separation and total distance for automated alignment verification. | Build a tool that places shapes, measures their separation, and supplies the metrics for downstream processing.
// AI Prompts: Show C# code that reads the Left and Top pixel values of two shapes in Aspose.Cells and computes their offset. | Provide an Aspose.Cells example that writes the Euclidean distance between two rectangle shapes to a cell. | Explain how to handle negative X or Y offsets when the second shape is positioned left or above the first shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds two rectangle shapes with absolute pixel positions, calculates the horizontal (X) and vertical (Y) pixel differences, optionally derives the Euclidean distance, writes these values to cells A1‑B3, and saves the file as ShapeOffsetResult.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add the first rectangle shape
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape shape1 = sheet.Shapes.AddRectangle(5, 2, 0, 0, 100, 50);
        // Set its absolute pixel position
        shape1.Left = 150; // horizontal offset in pixels from the left column
        shape1.Top = 200;  // vertical offset in pixels from the top row

        // Add the second rectangle shape
        Shape shape2 = sheet.Shapes.AddRectangle(10, 4, 0, 0, 120, 60);
        shape2.Left = 300;
        shape2.Top = 350;

        // Calculate horizontal (X) and vertical (Y) pixel offsets between the two shapes
        int offsetX = shape2.Left - shape1.Left; // positive if shape2 is to the right of shape1
        int offsetY = shape2.Top - shape1.Top;   // positive if shape2 is below shape1

        // Store the offsets in cells
        sheet.Cells["A1"].PutValue("OffsetX");
        sheet.Cells["B1"].PutValue(offsetX);
        sheet.Cells["A2"].PutValue("OffsetY");
        sheet.Cells["B2"].PutValue(offsetY);

        // Optionally store the Euclidean distance between the shapes
        double distance = Math.Sqrt(offsetX * offsetX + offsetY * offsetY);
        sheet.Cells["A3"].PutValue("Distance");
        sheet.Cells["B3"].PutValue(distance);

        // Save the workbook
        workbook.Save("ShapeOffsetResult.xlsx");
    }
}
