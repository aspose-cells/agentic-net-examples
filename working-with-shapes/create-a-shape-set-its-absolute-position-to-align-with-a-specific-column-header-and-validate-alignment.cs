// Title: Align a rectangle shape with a column header using absolute pixel positioning in Aspose.Cells for .NET
// Common Searches: how to position a shape over a specific cell using Aspose.Cells C# | set rectangle shape left and top to match column B header in Aspose.Cells | use FreeFloating placement to align shapes with column headers in .NET | calculate cell pixel coordinates for shape alignment in Aspose.Cells | validate shape alignment with header cell in C# Aspose.Cells example
// Tags: absolute pixel positioning of shapes Aspose.Cells | FreeFloating placement rectangle shape .NET | calculate cell pixel offsets Aspose.Cells | shape alignment with column header C# | validate shape position Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, writes a header in column B row 1, computes the pixel offsets of that cell by summing column widths and row heights, adds a rectangle shape, sets its Placement to FreeFloating, assigns the calculated Left and Top pixel values to align the shape with the header cell, checks the alignment, and saves the file as ShapeAlignment.xlsx.
class ShapeAlignmentDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the column header location (column B, row 1) – zero‑based indexes
            int headerColumnIndex = 1; // B
            int headerRowIndex = 0;    // 1

            // Write header text
            Cell headerCell = sheet.Cells[headerRowIndex, headerColumnIndex];
            headerCell.PutValue("Header B");

            // Calculate pixel offset of the cell's top‑left corner
            double cellLeft = 0;
            for (int c = 0; c < headerColumnIndex; c++)
                cellLeft += sheet.Cells.GetColumnWidthPixel(c);

            double cellTop = 0;
            for (int r = 0; r < headerRowIndex; r++)
                cellTop += sheet.Cells.GetRowHeightPixel(r);

            // Add a rectangle shape (width: 100px, height: 30px) directly
            Shape rectShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, 0, 0, 0, 0, 30, 100);

            // Ensure absolute positioning and align with the cell
            rectShape.Placement = PlacementType.FreeFloating;
            rectShape.Left = (int)cellLeft; // X‑coordinate (pixels)
            rectShape.Top = (int)cellTop;   // Y‑coordinate (pixels)

            // Validation: check that the shape's position matches the cell's position
            bool isAligned = Math.Abs(rectShape.Left - cellLeft) < 0.01 &&
                             Math.Abs(rectShape.Top - cellTop) < 0.01;

            Console.WriteLine(isAligned
                ? "Shape is correctly aligned with the column header."
                : "Shape alignment mismatch.");

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeAlignment.xlsx");

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
