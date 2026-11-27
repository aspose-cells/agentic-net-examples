// Title: Get a shape's absolute top‑left coordinates and align it to a target cell range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# retrieve shape absolute position in points | How to align a rectangle shape to a specific cell range with Aspose.Cells | Convert UpperDeltaX and UpperDeltaY pixel offsets to points in Aspose.Cells | Set shape anchor to the top‑left cell of a range using Aspose.Cells API | Resize Aspose.Cells shape to match the dimensions of a cell range in C#
// Tags: retrieve shape absolute coordinates Aspose.Cells | shape anchor cell range Aspose.Cells C# | convert shape pixel offsets to points Aspose.Cells | resize shape to match cell range Aspose.Cells | align shape to target range Aspose.Cells .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace ShapeAlignmentDemo
{
    // The example loads or creates a workbook, adds a rectangle shape, reads its UpperLeftRow/Column and UpperDeltaX/Y pixel offsets, converts the cell's top‑left corner and offsets to points to calculate the absolute position, then re‑anchors the shape to the top‑left cell of a target range (e.g., D10:E12), resets offsets, optionally resizes the shape to the range's width and height, verifies the new coordinates, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare a workbook (load if exists, otherwise create a new one)
                // -----------------------------------------------------------------
                string workbookPath = "input.xlsx";
                Workbook workbook;

                if (File.Exists(workbookPath))
                {
                    workbook = new Workbook(workbookPath);
                    Console.WriteLine($"Loaded workbook \"{workbookPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine("Created a new workbook.");
                }

                // -----------------------------------------------------------------
                // 2. Access the first worksheet
                // -----------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // -----------------------------------------------------------------
                // 3. Add a rectangle shape
                // -----------------------------------------------------------------
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 80);

                // Optional visual styling (commented out if API version does not support)
                // shape.Fill.ForeColor = Color.LightBlue;
                // shape.Line.Weight = 1.0;
                // shape.Line.DashStyle = Aspose.Cells.Drawing.LineDashStyle.Solid;

                Console.WriteLine("Added rectangle shape.");

                // -----------------------------------------------------------------
                // 4. Retrieve the absolute top‑left coordinates of the shape
                // -----------------------------------------------------------------
                int anchorRow = shape.UpperLeftRow;
                int anchorColumn = shape.UpperLeftColumn;
                int deltaX = shape.UpperDeltaX; // horizontal offset in pixels
                int deltaY = shape.UpperDeltaY; // vertical offset in pixels

                // Convert the cell's top‑left corner to points (1 point = 1/72 inch).
                AsposeRange anchorCellRange = sheet.Cells.CreateRange(anchorRow, anchorColumn, 1, 1);
                double cellLeftPoints = anchorCellRange.Left;
                double cellTopPoints = anchorCellRange.Top;

                // Convert pixel offsets to points (96 DPI is the default screen resolution used by Aspose.Cells)
                const double pointsPerPixel = 72.0 / 96.0;
                double offsetLeftPoints = deltaX * pointsPerPixel;
                double offsetTopPoints = deltaY * pointsPerPixel;

                double absoluteLeft = cellLeftPoints + offsetLeftPoints;
                double absoluteTop = cellTopPoints + offsetTopPoints;

                Console.WriteLine("=== Current Shape Position ===");
                Console.WriteLine($"Anchor cell: Row {anchorRow}, Column {anchorColumn}");
                Console.WriteLine($"Pixel offsets – X: {deltaX} px, Y: {deltaY} px");
                Console.WriteLine($"Absolute position – Left: {absoluteLeft:F2} pt, Top: {absoluteTop:F2} pt");

                // -----------------------------------------------------------------
                // 5. Align the shape to a target cell range (e.g., D10:E12)
                // -----------------------------------------------------------------
                string targetAddress = "D10:E12";
                AsposeRange targetRange = sheet.Cells.CreateRange(targetAddress);

                // Set the shape's anchor to the top‑left cell of the target range
                shape.UpperLeftRow = targetRange.FirstRow;
                shape.UpperLeftColumn = targetRange.FirstColumn;

                // Reset pixel offsets so the shape aligns exactly with the cell's top‑left corner
                shape.UpperDeltaX = 0;
                shape.UpperDeltaY = 0;

                // Optionally, resize the shape to match the target range size
                double rangeWidthPoints = targetRange.Width;   // points
                double rangeHeightPoints = targetRange.Height; // points
                int widthPixels = (int)Math.Round(rangeWidthPoints / pointsPerPixel);
                int heightPixels = (int)Math.Round(rangeHeightPoints / pointsPerPixel);
                shape.Width = widthPixels;
                shape.Height = heightPixels;

                // -----------------------------------------------------------------
                // 6. Verify the new absolute coordinates
                // -----------------------------------------------------------------
                int newAnchorRow = shape.UpperLeftRow;
                int newAnchorColumn = shape.UpperLeftColumn;
                int newDeltaX = shape.UpperDeltaX;
                int newDeltaY = shape.UpperDeltaY;

                AsposeRange newAnchorCellRange = sheet.Cells.CreateRange(newAnchorRow, newAnchorColumn, 1, 1);
                double newCellLeftPoints = newAnchorCellRange.Left;
                double newCellTopPoints = newAnchorCellRange.Top;

                double newAbsoluteLeft = newCellLeftPoints + newDeltaX * pointsPerPixel;
                double newAbsoluteTop = newCellTopPoints + newDeltaY * pointsPerPixel;

                Console.WriteLine("\n=== Shape Position After Alignment ===");
                Console.WriteLine($"Target range: {targetAddress}");
                Console.WriteLine($"New anchor cell: Row {newAnchorRow}, Column {newAnchorColumn}");
                Console.WriteLine($"New absolute position – Left: {newAbsoluteLeft:F2} pt, Top: {newAbsoluteTop:F2} pt");

                // -----------------------------------------------------------------
                // 7. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "ShapeAlignmentResult.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nWorkbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
