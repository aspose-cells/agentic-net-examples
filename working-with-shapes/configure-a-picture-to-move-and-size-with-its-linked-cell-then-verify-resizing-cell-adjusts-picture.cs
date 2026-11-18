// Title: How to anchor a picture to a cell and make it move and resize with the cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set picture placement to MoveAndSize so it follows cell resizing | C# example to link an image to a worksheet cell and have it resize when column width changes | Verify picture dimensions after adjusting row height and column width with Aspose.Cells | How to use PlacementType.MoveAndSize for pictures anchored to a specific cell in Aspose.Cells | Aspose.Cells picture anchored to B2 automatically updates size on cell resize
// Tags: PlacementType.MoveAndSize picture anchoring | link picture to cell B2 Aspose.Cells | automatic picture resizing on column width change | C# picture size verification after cell resize | Aspose.Cells picture anchored to worksheet cell

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a workbook, adds a PNG picture anchored to cell B2, sets its Placement to MoveAndSize, records the original dimensions, changes the column width and row height of the linked cell, then reads the updated picture size to demonstrate automatic resizing, and finally saves the file as PictureLinkTest.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // A tiny 1x1 PNG image (transparent) encoded in Base64
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            // Add picture anchored to cell B2 (row index 1, column index 1)
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                // Pictures.Add returns the index of the newly added picture
                int pictureIndex = sheet.Pictures.Add(1, 1, ms);
                Picture picture = sheet.Pictures[pictureIndex];
                // Configure the picture to move and size with its linked cell
                picture.Placement = PlacementType.MoveAndSize;
            }

            // Capture original picture dimensions
            Picture pic = sheet.Pictures[0];
            double originalWidth = pic.Width;
            double originalHeight = pic.Height;

            // Resize the linked cell (B2) by changing column width and row height
            sheet.Cells.SetColumnWidth(1, 30); // Column width in characters
            sheet.Cells.SetRowHeight(1, 60);   // Row height in points

            // After resizing the cell, the picture should have updated dimensions
            double newWidth = pic.Width;
            double newHeight = pic.Height;

            // Output the results to verify the behavior
            Console.WriteLine($"Original picture size: {originalWidth} x {originalHeight}");
            Console.WriteLine($"After cell resize: {newWidth} x {newHeight}");

            // Save the workbook (optional, for manual inspection)
            string outputPath = "PictureLinkTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
