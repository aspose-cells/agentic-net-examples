// Title: Add a rectangle shape and position it at a specific column using absolute coordinates in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to the first worksheet, switches it to free‑floating placement, and sets its Left property to the cumulative width of columns before column C (converted to points). | Write a method that sums the pixel widths of all columns preceding a given zero‑based column index, converts the total to points, and assigns the result to a shape's Left coordinate in an Aspose.Cells workbook. | Create a console application that inserts a shape, positions it with absolute coordinates based on a column index, prints the shape's final Left and Top values, and saves the workbook to a file.
// Common Searches: Aspose.Cells C# set shape left position based on column width | How to calculate absolute position of a shape using column index in Aspose.Cells | Free floating shape placement with column pixel conversion Aspose.Cells .NET | C# Aspose.Cells get column width in pixels and convert to points for shape positioning | Place rectangle shape at column C in Excel using Aspose.Cells API
// Tags: add rectangle shape Aspose.Cells C# | freefloating shape placement Aspose.Cells | calculate column width pixels Aspose.Cells | convert pixels to points Aspose.Cells | absolute shape coordinates column index

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, computes the left offset in points by summing pixel widths of columns before a target column, adds a rectangle shape anchored at cell A1, switches it to free‑floating placement, applies the calculated Left and Top values, outputs the final coordinates, and saves the file as ShapePlacementDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target column index (0‑based). Example: column C => index 2
            int targetColumnIndex = 2;

            // Calculate the absolute left position in points.
            // Sum the pixel widths of all columns before the target column,
            // then convert pixels to points (1 pixel ≈ 0.75 point at 96 DPI).
            double totalPixels = 0;
            for (int col = 0; col < targetColumnIndex; col++)
            {
                int colPixels = sheet.Cells.GetColumnWidthPixel(col);
                totalPixels += colPixels;
            }
            double leftPosition = totalPixels * 0.75; // points

            // Define top position, width and height for the shape (in points)
            double topPosition = 0.0;
            int shapeWidth = 100;   // width in points
            int shapeHeight = 50;   // height in points

            // Add a rectangle shape anchored to cell (0,0) with zero offsets
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                0,          // upper left row
                0,          // upper left column
                0,          // upper left row offset (pixels)
                0,          // upper left column offset (pixels)
                shapeHeight,
                shapeWidth);

            // Switch to absolute positioning and set exact coordinates
            shape.Placement = PlacementType.FreeFloating;
            shape.Left = (int)leftPosition;   // cast if Left expects int
            shape.Top = (int)topPosition;     // cast if Top expects int
            shape.Text = "Test Shape";

            // Verify placement by outputting the shape's coordinates
            Console.WriteLine($"Shape placed at Left = {shape.Left} points, Top = {shape.Top} points.");

            // Determine output file path and ensure the directory exists
            string outputPath = "ShapePlacementDemo.xlsx";
            try
            {
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
