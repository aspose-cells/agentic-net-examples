// Title: Add a rectangle shape over a merged cell range and confirm its exact position with Aspose.Cells for .NET
// Common Searches: how to add a shape that covers merged cells in Aspose.Cells C# | Aspose.Cells C# place rectangle over merged range B2:D4 | verify shape dimensions match merged cell area using Aspose.Cells | calculate pixel width and height of merged cells for shape placement Aspose.Cells | set shape UpperLeftRow UpperLeftColumn to align with merged cells Aspose.Cells
// Tags: rectangle shape placement on merged range Aspose.Cells | absolute shape coordinates Excel C# | merged cell pixel size computation Aspose.Cells | shape alignment check Aspose.Cells | Aspose.Cells shape sizing for merged cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells B2:D4, computes the combined pixel width and height of the merged area, adds a rectangle shape sized to cover that area, sets the shape's UpperLeftRow, UpperLeftColumn, Top, Left, Width, and Height to match the merged cells, validates the alignment within a small tolerance, and saves the file as MergedShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged cell range (B2:D4)
            int startRow = 1;      // B2 (zero‑based)
            int startColumn = 1;   // B
            int endRow = 3;        // D4 row index
            int endColumn = 3;     // D column index

            // Merge the cells
            sheet.Cells.Merge(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);

            // Calculate total width and height of the merged area in pixels
            int totalWidth = 0;
            for (int col = startColumn; col <= endColumn; col++)
                totalWidth += sheet.Cells.GetColumnWidthPixel(col);

            int totalHeight = 0;
            for (int row = startRow; row <= endRow; row++)
                totalHeight += sheet.Cells.GetRowHeightPixel(row);

            // Add a rectangle shape that will cover the merged area
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                startRow,          // upper left row
                startColumn,       // upper left column
                0,                 // top offset (pixels)
                0,                 // left offset (pixels)
                totalHeight,       // height (pixels)
                totalWidth);       // width (pixels)

            // Position the shape to align exactly with the merged cells
            shape.UpperLeftRow = startRow;
            shape.UpperLeftColumn = startColumn;
            shape.Top = 0;
            shape.Left = 0;
            shape.Width = totalWidth;
            shape.Height = totalHeight;

            // Verify alignment (allow a small tolerance)
            bool aligns =
                Math.Abs(shape.Left - sheet.Cells.GetColumnWidthPixel(startColumn) * 0) < 0.1 &&
                Math.Abs(shape.Top - sheet.Cells.GetRowHeightPixel(startRow) * 0) < 0.1 &&
                Math.Abs(shape.Width - totalWidth) < 0.1 &&
                Math.Abs(shape.Height - totalHeight) < 0.1;

            Console.WriteLine("Shape aligns with merged area: " + aligns);

            // Save the workbook
            string outputPath = "MergedShapeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
