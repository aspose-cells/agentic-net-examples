// Title: Create a dark gray line shape in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add line shape with custom color to worksheet | set line shape color to dark gray using Aspose.Cells .NET | how to draw a horizontal line shape in Excel with Aspose.Cells C# | Aspose.Cells line shape placement MoveAndSize example | change line thickness and dash style in Aspose.Cells C#
// Tags: add line shape Aspose.Cells C# | set line color Aspose.Cells | line shape formatting Aspose.Cells | line shape placement MoveAndSize Aspose.Cells | horizontal line shape Excel Aspose.Cells | line weight dash style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddLineShapeDemo
{
    // The sample creates a new workbook, adds a horizontal LineShape at a specific cell location, sets its color to dark gray, adjusts line weight and dash style, configures the shape to move and size with cells, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the ShapeCollection from the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Define the position and size of the line (in pixels)
            // topRow, top, leftColumn, left, height, width
            // Here we create a horizontal line 200 pixels wide and 0 pixels high (height is ignored for a line)
            int topRow = 2;      // Row index where the line starts
            int top = 0;         // Vertical offset within the row (pixels)
            int leftColumn = 2;  // Column index where the line starts
            int left = 0;        // Horizontal offset within the column (pixels)
            int height = 0;      // Height of the line (0 for a straight line)
            int width = 200;     // Length of the line (pixels)

            // Add the line shape
            LineShape line = shapes.AddLine(topRow, top, leftColumn, left, height, width);

            // Set the line's appearance
            // Use the modern Line property (not the obsolete LineFormat)
            line.Line.SolidFill.Color = Color.DarkGray;   // Set line color to dark gray
            line.Line.Weight = 2;                         // Optional: set line thickness
            line.Line.DashStyle = MsoLineDashStyle.Solid; // Optional: solid line

            // Ensure the shape moves and sizes with cells (optional)
            line.Placement = PlacementType.MoveAndSize;

            // Save the workbook to a file
            string outputPath = "LineShapeDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
