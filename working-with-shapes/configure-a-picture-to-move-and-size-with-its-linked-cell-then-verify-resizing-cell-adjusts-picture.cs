using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class PictureMoveAndSizeWithCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a picture to the worksheet at cell A1 (row 0, column 0)
        // Ensure you have an image file named "sample.png" in the executable directory
        int picIndex = sheet.Pictures.Add(0, 0, "sample.png");
        Picture picture = sheet.Pictures[picIndex];

        // Link the picture to cell B2 and set it to move and size with the cell
        picture.LinkedCell = "B2";
        picture.Placement = PlacementType.MoveAndSize;   // Move and size with cells

        // Record original picture dimensions
        double originalHeight = picture.Height;
        double originalWidth = picture.Width;
        Console.WriteLine($"Original picture size: Height={originalHeight}, Width={originalWidth}");

        // Resize the linked cell (B2)
        // Increase row height and column width to see the effect on the picture
        sheet.Cells.SetRowHeight(1, 60);      // Row index 1 corresponds to row 2 (B2)
        sheet.Cells.SetColumnWidth(1, 25);   // Column index 1 corresponds to column B

        // After resizing, retrieve the picture dimensions again
        double newHeight = picture.Height;
        double newWidth = picture.Width;
        Console.WriteLine($"After resizing cell B2: Picture size: Height={newHeight}, Width={newWidth}");

        // Verify that the picture has been resized (simple check)
        bool resized = Math.Abs(newHeight - originalHeight) > 0.1 || Math.Abs(newWidth - originalWidth) > 0.1;
        Console.WriteLine($"Picture resized with cell: {resized}");

        // Save the workbook to verify the result visually if needed
        workbook.Save("PictureMoveAndSizeWithCell.xlsx", SaveFormat.Xlsx);
    }
}