// Title: Resize a horizontal line shape to span cells B2:E2 and set its thickness in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to draw a line that covers cells B2 to E2 | C# set line shape width based on column pixel widths in Aspose.Cells | adjust line thickness in Aspose.Cells .NET workbook | make line shape free floating so it doesn't move with cells in Aspose.Cells | calculate total column width in pixels for a range using Aspose.Cells | add horizontal line shape across multiple columns in Aspose.Cells C#
// Tags: horizontal line shape spanning multiple columns Aspose.Cells | set line weight points Aspose.Cells C# | shape placement type free floating Aspose.Cells | column width pixel calculation Aspose.Cells | add line shape to worksheet Aspose.Cells | adjust line shape dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, computes the combined pixel width of columns B through E in row 2, adds a horizontal LineShape that starts at column B and ends at column E, sets the shape to free‑floating placement, applies a 3‑point solid line weight with a dark blue color, and saves the file as 'LineResizeDemo.xlsx'.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Determine the pixel coordinates that correspond to cells B2:E2
        // ------------------------------------------------------------
        // Rows and columns in Aspose.Cells are zero‑based.
        // B2  -> column index 1, row index 1
        // E2  -> column index 4, row index 1
        int startRow = 1;          // row 2 in Excel
        int startColumn = 1;       // column B
        int endColumn = 4;         // column E

        // Calculate the total width (in pixels) that spans from the left edge of B
        // to the right edge of E.
        double totalWidthPixels = 0;
        for (int col = startColumn; col <= endColumn; col++)
        {
            // GetColumnWidthPixel returns the width of a column in pixels.
            totalWidthPixels += worksheet.Cells.GetColumnWidthPixel(col);
        }

        // Height of a horizontal line is effectively zero; we use 0.
        // The line will be placed at the top of the row (offset = 0).
        int lineHeightPixels = 0;

        // ------------------------------------------------------------
        // Add the line shape
        // ------------------------------------------------------------
        ShapeCollection shapes = worksheet.Shapes;

        // AddLine(topRow, top, leftColumn, left, height, width)
        // top and left offsets are set to 0 (no extra pixel offset).
        LineShape line = shapes.AddLine(
            topRow: startRow,
            top: 0,
            leftColumn: startColumn,
            left: 0,
            height: lineHeightPixels,
            width: (int)totalWidthPixels);

        // ------------------------------------------------------------
        // Adjust line appearance
        // ------------------------------------------------------------
        // Make the line free‑floating so it does not move/size with cells.
        line.Placement = PlacementType.FreeFloating;

        // Set line thickness (weight) – value is in points.
        line.Line.Weight = 3; // 3‑point thickness

        // Ensure the line style is solid.
        line.Line.DashStyle = MsoLineDashStyle.Solid;

        // Optional: set line color (e.g., dark blue)
        line.Line.SolidFill.Color = System.Drawing.Color.DarkBlue;

        // ------------------------------------------------------------
        // Save the workbook
        // ------------------------------------------------------------
        string outputPath = "LineResizeDemo.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Line spans B2:E2 with thickness 3 points.");
    }
}
