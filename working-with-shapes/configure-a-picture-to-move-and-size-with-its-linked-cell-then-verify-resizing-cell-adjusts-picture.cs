// Title: How to set a picture to move and size with its linked cell and verify the effect after resizing the cell using Aspose.Cells for .NET
// AI Prompts: Create a MemoryStream from PNG bytes, insert the image as a picture anchored at cell B2, and set its Placement so the picture follows the cell when the cell moves or resizes. | Capture the picture's initial Width and Height, change the row height and column width of the linked cell, then read and output the picture's new dimensions.
// Common Searches: Aspose.Cells picture placement move and size with linked cell example | C# resize Excel cell and automatically adjust anchored image size | Add in‑memory PNG to worksheet and bind it to a cell using Aspose.Cells | Verify image size after adjusting row height and column width in .NET
// Tags: picture placement MoveAndSize Aspose.Cells | anchor image to worksheet cell .NET | cell dimension change updates picture size | insert PNG via MemoryStream into Excel worksheet | measure picture dimensions after cell resize

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, inserts a PNG image into cell B2 via a MemoryStream, sets the picture's Placement to MoveAndSize so it moves and resizes with the cell, records the original picture size, enlarges the row height and column width of the linked cell, then reads and prints the updated picture dimensions before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some text in the cell that will be linked to the picture (B2)
            sheet.Cells["B2"].PutValue("Linked Cell");

            // A minimal 1x1 PNG image (transparent) encoded in Base64
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            // Add the picture anchored to cell B2 (row index 1, column index 1)
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                // Pictures.Add returns the index of the newly added picture
                int pictureIndex = sheet.Pictures.Add(1, 1, ms);
                Picture picture = sheet.Pictures[pictureIndex];
                // Configure the picture to move and size with its linked cell
                picture.Placement = PlacementType.MoveAndSize;
            }

            // Capture the picture's original dimensions
            Picture linkedPic = sheet.Pictures[0];
            double originalWidth = linkedPic.Width;
            double originalHeight = linkedPic.Height;

            // Resize the linked cell (B2) by changing row height and column width
            sheet.Cells.SetRowHeight(1, 100);   // Row 2 height = 100 points
            sheet.Cells.SetColumnWidth(1, 30); // Column B width = 30 characters

            // After resizing the cell, the picture should have updated dimensions
            double newWidth = linkedPic.Width;
            double newHeight = linkedPic.Height;

            // Output verification results
            Console.WriteLine($"Original picture size: {originalWidth} x {originalHeight}");
            Console.WriteLine($"After cell resize picture size: {newWidth} x {newHeight}");

            // Save the workbook (optional, for manual inspection)
            string outputPath = "PictureLinkDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
