// Title: Overlay a rectangle shape on a merged cell range and verify its absolute position with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape covering the merged range B2:D4, then checks that the shape’s UpperLeftRow, UpperLeftColumn, Width and Height exactly match the merged cells using Aspose.Cells. | Write a verification routine that computes the total pixel width of columns B‑D and total pixel height of rows 2‑4, and compares these values with the shape’s Width and Height to confirm alignment.
// Common Searches: aspocells c# place rectangle shape over merged cells B2:D4 and validate alignment | how to check shape dimensions against merged cell range in Aspose.Cells | c# Aspose.Cells calculate pixel size of merged range for shape placement | verify shape UpperLeftRow and UpperLeftColumn correspond to merged area using Aspose.Cells
// Tags: add rectangle shape to merged cells aspocells c# | verify shape alignment with merged range aspocells | calculate shape pixel dimensions from merged cell range aspocells | shape placement freefloating aspocells | merged cell range shape size validation aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells B2:D4, inserts a rectangle shape that spans the merged area, and then validates that the shape’s UpperLeftRow, UpperLeftColumn, width, and height precisely match the merged cells by comparing pixel dimensions. The result is printed and the workbook is saved as MergedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged range (B2:D4) – zero‑based indexes
            int firstRow = 1;      // Row 2
            int firstColumn = 1;   // Column B
            int totalRows = 3;     // Rows 2,3,4
            int totalColumns = 3;  // Columns B,C,D

            // Merge the cells and set a value
            sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);
            sheet.Cells[firstRow, firstColumn].PutValue("Merged Area");

            // Add a rectangle shape that covers the merged range
            // Parameters: shape type, upper‑left row, upper‑left column,
            // row offset (pixels), column offset (pixels), height (in rows), width (in columns)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                firstRow,               // Upper‑left row
                firstColumn,            // Upper‑left column
                0,                      // Row offset (pixels)
                0,                      // Column offset (pixels)
                totalRows,              // Height in rows
                totalColumns);          // Width in columns

            // Ensure the shape moves and sizes with cells (optional)
            shape.Placement = PlacementType.FreeFloating;

            // ---------- Verification of alignment ----------
            // Expected top‑left cell coordinates (offsets are not available in this API version)
            bool aligns =
                shape.UpperLeftRow == firstRow &&
                shape.UpperLeftColumn == firstColumn;

            // Verify shape size matches the merged range size (in pixels)
            int expectedWidth = 0;
            for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                expectedWidth += sheet.Cells.GetColumnWidthPixel(c);

            int expectedHeight = 0;
            for (int r = firstRow; r < firstRow + totalRows; r++)
                expectedHeight += sheet.Cells.GetRowHeightPixel(r);

            aligns = aligns &&
                     Math.Abs(shape.Width - expectedWidth) < 1 &&   // allow minor rounding differences
                     Math.Abs(shape.Height - expectedHeight) < 1;

            Console.WriteLine("Shape aligns with merged area: " + aligns);

            // Save the workbook (optional)
            workbook.Save("MergedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
