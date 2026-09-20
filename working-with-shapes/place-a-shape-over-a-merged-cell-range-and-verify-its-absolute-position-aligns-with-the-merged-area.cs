// Title: Overlay a rectangle shape on a merged cell range and validate its absolute coordinates using Aspose.Cells for .NET
// AI Prompts: Calculate the pixel left, top, width, and height of a merged range (e.g., B2:D4) and assign those values to a rectangle shape's Left, Top, Width, and Height properties with Aspose.Cells. | Compare the shape's absolute position properties to the computed merged‑range boundaries and output a boolean indicating whether the shape aligns perfectly.
// Common Searches: Aspose.Cells .NET how to place a shape exactly over merged cells | C# get pixel dimensions of a merged cell range in Aspose.Cells | verify shape alignment with merged area using Aspose.Cells API | set rectangle shape size to match merged cells B2:D4 Aspose.Cells | calculate left and top offsets for shapes over merged cells in C#
// Tags: overlay shape on merged cells Aspose.Cells | merged range pixel size calculation .NET | shape absolute positioning Aspose.Cells | rectangle shape size matching merged area | shape alignment check C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells B2:D4, adds a rectangle shape, computes the merged range's pixel boundaries, sets the shape's Left, Top, Width, and Height to exactly cover the merged area, verifies the alignment with a boolean check, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells B2:D4 (zero‑based indices: rows 1‑3, columns 1‑3)
            int startRow = 1;      // Row 2
            int startColumn = 1;   // Column B
            int totalRows = 3;     // 3 rows (2,3,4)
            int totalColumns = 3;  // 3 columns (B,C,D)
            sheet.Cells.Merge(startRow, startColumn, totalRows, totalColumns);

            // Add a rectangle shape (initial size will be overridden)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, 100, 100);

            // Calculate pixel boundaries of the merged range
            int left = 0;
            for (int c = 0; c < startColumn; c++)
                left += sheet.Cells.GetColumnWidthPixel(c);

            int top = 0;
            for (int r = 0; r < startRow; r++)
                top += sheet.Cells.GetRowHeightPixel(r);

            int width = 0;
            for (int c = startColumn; c < startColumn + totalColumns; c++)
                width += sheet.Cells.GetColumnWidthPixel(c);

            int height = 0;
            for (int r = startRow; r < startRow + totalRows; r++)
                height += sheet.Cells.GetRowHeightPixel(r);

            // Position the shape to exactly cover the merged area
            shape.Left = left;
            shape.Top = top;
            shape.Width = width;
            shape.Height = height;

            // Verify that the shape's absolute position matches the merged area
            bool aligns = Math.Abs(shape.Left - left) < 0.1 &&
                          Math.Abs(shape.Top - top) < 0.1 &&
                          Math.Abs(shape.Width - width) < 0.1 &&
                          Math.Abs(shape.Height - height) < 0.1;

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
