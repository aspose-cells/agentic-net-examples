// Title: Create a rectangle shape with custom width and height and anchor it to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape with custom dimensions to a specific cell range using Aspose.Cells C# | Aspose.Cells C# set shape width and height and anchor to range A2:B4 | C# Aspose.Cells place shape with PlacementType.MoveAndSize on a worksheet | example of adding a rectangle shape to a workbook and resizing it in points with Aspose.Cells | Aspose.Cells add shape anchored to cells and define size programmatically
// Tags: Aspose.Cells rectangle shape C# | set shape width height Aspose.Cells | anchor shape to cell range Aspose.Cells | PlacementType.MoveAndSize usage Aspose.Cells | add shape to worksheet Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    // The example creates a new workbook, adds a rectangle shape anchored to the cell range A2:B4 on the first worksheet, sets its width to 200 points and height to 100 points, applies the MoveAndSize placement, and saves the file as OutputWithShape.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell range that will anchor the shape (e.g., A2:B4)
                int upperLeftRow = 1;      // Row index for A2 (zero‑based)
                int upperLeftColumn = 0;   // Column index for A2

                // Add a rectangle shape anchored to the upper‑left cell of the range
                // (top, left, height, width are initially set to 0; we'll set custom size later)
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow,
                    upperLeftColumn,
                    0,    // top offset in pixels
                    0,    // left offset in pixels
                    0,    // initial height
                    0);   // initial width

                // Set custom dimensions (width and height in points)
                shape.Width = 200;   // custom width
                shape.Height = 100;  // custom height

                // Keep the shape anchored to the defined cell range
                shape.Placement = PlacementType.MoveAndSize;

                // Define output file path
                string outputPath = "OutputWithShape.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
