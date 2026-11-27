// Title: Link a rectangle shape to cell H5 and enable move‑and‑size behavior with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape to a specific cell and set move and size placement | Set PlacementType.MoveAndSize for a rectangle shape linked to H5 in Aspose.Cells | C# example of a shape following cell resizing using Aspose.Cells | How to bind a shape to a cell and keep it synchronized with cell size in Aspose.Cells | Aspose.Cells LinkedCell property usage with shape placement options
// Tags: Aspose.Cells shape linked cell | PlacementType.MoveAndSize rectangle | C# Aspose.Cells shape placement | link shape to cell H5 Aspose.Cells | shape move with cell resizing .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeLinkedCellDemo
{
    // Creates a new workbook, adds a rectangle shape, links it to cell H5, sets the shape's Placement to MoveAndSize so it moves and resizes with the cell, applies a light‑blue fill and solid line, and saves the workbook as ShapeLinkedCellDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (you can also load an existing file if needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, top offset (pixels), upper left column, left offset (pixels), height (pixels), width (pixels)
                int upperLeftRow = 2;      // Row index (0‑based)
                int top = 0;               // Pixel offset from the top of the row
                int upperLeftColumn = 2;   // Column index (0‑based)
                int left = 0;              // Pixel offset from the left of the column
                int height = 100;          // Height in pixels
                int width = 200;           // Width in pixels

                Shape rectangle = worksheet.Shapes.AddRectangle(upperLeftRow, top, upperLeftColumn, left, height, width);

                // Link the shape to cell H5
                // The LinkedCell property expects an A1‑style reference (absolute or relative)
                rectangle.LinkedCell = "H5";

                // Set the shape to move and resize with the linked cell
                // PlacementType.MoveAndSize makes the shape follow cell size changes
                rectangle.Placement = PlacementType.MoveAndSize;

                // Optional: give the shape a visible fill and line so you can see it in the saved file
                rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
                rectangle.Line.Weight = 2;
                rectangle.Line.DashStyle = MsoLineDashStyle.Solid;

                // Save the workbook to a file
                string outputPath = "ShapeLinkedCellDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The rectangle shape is linked to cell H5 and will move/size with the cell.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
