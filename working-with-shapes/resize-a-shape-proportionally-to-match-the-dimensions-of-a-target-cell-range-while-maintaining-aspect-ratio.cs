// Title: Proportionally Resize an Aspose.Cells Shape to Fit a Cell Range (C#)
// Description: Demonstrates how to calculate the pixel width and height of a target cell range, compute a uniform scaling factor that preserves the shape's aspect ratio, resize the shape accordingly, and move it into the specified range using Aspose.Cells for .NET.
// Keywords: Aspose.Cells resize shape | scale shape proportionally .NET | maintain aspect ratio Aspose.Cells | move shape to cell range | calculate cell width pixel Aspose.Cells | C# shape scaling Aspose.Cells | fit shape into merged cells
// Common Searches: How to resize a shape to a cell range in Aspose.Cells C# | Preserve aspect ratio when scaling shapes with Aspose.Cells | Move rectangle shape to specific cells Aspose.Cells | Get pixel dimensions of a range in Aspose.Cells | Proportional shape scaling example Aspose.Cells
// Developer Intent: Resize a shape so it fits inside a given cell range while keeping its original aspect ratio.
// Use Cases: Insert a company logo into a header block without distortion. | Adjust a placeholder shape for a chart to match a dynamic report area. | Resize a textbox to occupy a merged cell region while preserving text layout.
// AI Prompts: Write C# code using Aspose.Cells to proportionally resize any shape to a target cell range and then place it in that range. | Explain the steps to compute a uniform scaling factor based on the pixel width and height of a cell range in Aspose.Cells. | Show how to move a resized shape to a specific range after scaling it proportionally with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to calculate the pixel width and height of a target cell range, compute a uniform scaling factor that preserves the shape's aspect ratio, resize the shape accordingly, and move it into the specified range using Aspose.Cells for .NET.
    class ResizeShapeProportionally
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (you can also add a picture, textbox, etc.)
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 150, 100);

                // Define the target cell range to which the shape should be resized
                // Example: range from row 5, column 2 to row 7, column 4 (zero‑based indices)
                int targetTopRow = 5;
                int targetLeftColumn = 2;
                int targetBottomRow = 7;
                int targetRightColumn = 4;

                // Calculate the total width of the target range in pixels
                double targetWidth = 0;
                for (int col = targetLeftColumn; col <= targetRightColumn; col++)
                {
                    targetWidth += sheet.Cells.GetColumnWidthPixel(col);
                }

                // Calculate the total height of the target range in pixels
                double targetHeight = 0;
                for (int row = targetTopRow; row <= targetBottomRow; row++)
                {
                    targetHeight += sheet.Cells.GetRowHeightPixel(row);
                }

                // Determine the scaling factor that preserves the shape's aspect ratio
                double widthScale = targetWidth / shape.Width;
                double heightScale = targetHeight / shape.Height;
                double uniformScale = Math.Min(widthScale, heightScale); // keep aspect ratio

                // Apply the uniform scaling to both dimensions (explicit cast for safety)
                shape.Width = (int)(shape.Width * uniformScale);
                shape.Height = (int)(shape.Height * uniformScale);

                // Move the resized shape to the target range
                shape.MoveToRange(targetTopRow, targetLeftColumn, targetBottomRow, targetRightColumn);

                // Prepare output path and ensure directory exists
                string outputPath = "ResizeShapeProportionally.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
