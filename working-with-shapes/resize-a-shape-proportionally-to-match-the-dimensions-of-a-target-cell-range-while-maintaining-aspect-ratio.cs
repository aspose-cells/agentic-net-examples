// Title: C# – Proportionally Resize an Aspose.Cells Shape to Fit a Cell Range
// Description: Creates a workbook, adds a shape, locks its aspect ratio, computes the target range size in points, scales the shape while preserving proportions, moves it to the range, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape resize | C# shape aspect ratio | fit shape to cell range | convert column width to points | move shape to range | Aspose.Cells .NET example | ResizeShapeProportionally
// Common Searches: How to resize a shape proportionally in Aspose.Cells C# | Fit shape inside specific rows and columns Aspose.Cells | Convert Excel column width and row height to points Aspose.Cells | Programmatically move a shape to a cell range Aspose.Cells | Lock aspect ratio for shapes in Aspose.Cells
// Developer Intent: Resize a shape so it fits within a defined cell range while preserving its original aspect ratio.
// Use Cases: Insert a company logo into a header area without distortion. | Adjust a chart placeholder to occupy a dynamic table range while keeping proportions. | Automatically scale imported images to match a report section defined by cells. | Resize comment boxes or callouts to fit within merged cells.
// AI Prompts: Generate C# code that uses Aspose.Cells to proportionally resize any shape to a given cell range and lock its aspect ratio. | Explain how to convert Excel column widths and row heights from pixels to points for accurate shape sizing in Aspose.Cells. | Show how to move a resized shape to a target range and keep it centered using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a shape, locks its aspect ratio, computes the target range size in points, scales the shape while preserving proportions, moves it to the range, and saves the file using Aspose.Cells for .NET.
    public class ResizeShapeProportionally
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (you can use any shape type)
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 100);

                // Optional: lock aspect ratio so manual resizing respects it
                shape.IsAspectRatioLocked = true;

                // Define the target range where the shape should fit (zero‑based indices)
                int topRow = 5;
                int leftColumn = 3;
                int bottomRow = 9;
                int rightColumn = 7;

                // ---------- Calculate target width and height in points ----------
                // Sum column widths (in pixels) for the columns that belong to the range
                double targetWidthPixels = 0;
                for (int col = leftColumn; col <= rightColumn; col++)
                {
                    targetWidthPixels += worksheet.Cells.GetColumnWidthPixel(col);
                }

                // Sum row heights (in pixels) for the rows that belong to the range
                double targetHeightPixels = 0;
                for (int row = topRow; row <= bottomRow; row++)
                {
                    targetHeightPixels += worksheet.Cells.GetRowHeightPixel(row);
                }

                // Convert pixels to points (1 pixel = 0.75 point at 96 DPI)
                const double pixelToPoint = 0.75;
                double targetWidthPoints = targetWidthPixels * pixelToPoint;
                double targetHeightPoints = targetHeightPixels * pixelToPoint;

                // ---------- Determine scaling factor while preserving aspect ratio ----------
                double originalWidth = shape.Width;   // current width in points
                double originalHeight = shape.Height; // current height in points

                // Compute the scale that fits both width and height inside the target area
                double widthScale = targetWidthPoints / originalWidth;
                double heightScale = targetHeightPoints / originalHeight;
                double scale = Math.Min(widthScale, heightScale); // keep aspect ratio

                // Apply the new size (cast to int if the API expects integer values)
                shape.Width = (int)(originalWidth * scale);
                shape.Height = (int)(originalHeight * scale);

                // ---------- Move the shape to the target range ----------
                shape.MoveToRange(topRow, leftColumn, bottomRow, rightColumn);

                // Save the workbook
                string outputPath = "ResizeShapeProportionally.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log any runtime errors that occur during processing
                Console.WriteLine($"Run error: {ex.Message}");
                throw; // Re‑throw to be caught by Main if needed
            }
        }
    }
}
