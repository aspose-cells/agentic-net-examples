// Title: Insert a PNG picture and proportionally scale it to fit a target cell range using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells insert image into B2:D10 while preserving aspect ratio | scale picture to match cell range dimensions Aspose.Cells .NET | center an inserted picture inside a worksheet range with Aspose.Cells | calculate pixel width and height of a cell area for image scaling Aspose.Cells | fit PNG into multiple cells without distortion using Aspose.Cells C#
// Tags: picture insertion into defined cell area Aspose.Cells | aspect‑ratio‑preserving scaling of PNG in worksheet | cell range pixel size calculation for image placement | centered picture positioning within multiple cells | dynamic image resizing based on worksheet dimensions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, defines a target cell range (e.g., B2:D10), inserts a PNG image, computes the total pixel width and height of the range, determines a scaling factor that maintains the image's aspect ratio, resizes the picture accordingly, optionally centers it within the range, and saves the result as Result.xlsx while handling missing files and other exceptions.
class InsertAndScalePicture
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range where the picture should fit (e.g., B2:D10)
            string startCell = "B2";
            string endCell = "D10";
            CellArea targetArea = CellArea.CreateCellArea(startCell, endCell);

            // Path to the picture file to be inserted
            string picturePath = "sample.png";

            // Ensure the picture file exists
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Insert the picture at the top‑left cell of the target range
            int pictureIndex = sheet.Pictures.Add(targetArea.StartRow, targetArea.StartColumn, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // ------------------------------------------------------------
            // Calculate the total pixel width and height of the target range
            // ------------------------------------------------------------
            double targetWidth = 0;
            for (int col = targetArea.StartColumn; col <= targetArea.EndColumn; col++)
                targetWidth += sheet.Cells.GetColumnWidthPixel(col);

            double targetHeight = 0;
            for (int row = targetArea.StartRow; row <= targetArea.EndRow; row++)
                targetHeight += sheet.Cells.GetRowHeightPixel(row);

            // ------------------------------------------------------------
            // Get original picture dimensions (in pixels)
            // ------------------------------------------------------------
            double originalWidth = picture.Width;
            double originalHeight = picture.Height;

            // ------------------------------------------------------------
            // Compute scaling factor to fit proportionally within the target range
            // ------------------------------------------------------------
            double widthScale = targetWidth / originalWidth;
            double heightScale = targetHeight / originalHeight;
            double scale = Math.Min(widthScale, heightScale); // preserve aspect ratio

            // Apply scaling (cast to int because Picture properties are int)
            picture.Width = (int)(originalWidth * scale);
            picture.Height = (int)(originalHeight * scale);

            // ------------------------------------------------------------
            // Optionally, center the picture inside the target range
            // ------------------------------------------------------------
            double excessWidth = targetWidth - picture.Width;
            double excessHeight = targetHeight - picture.Height;

            // Compute left offset: sum of column widths before start column + half of excess width
            double leftOffset = 0;
            for (int col = 0; col < targetArea.StartColumn; col++)
                leftOffset += sheet.Cells.GetColumnWidthPixel(col);
            leftOffset += excessWidth / 2;

            // Compute top offset: sum of row heights before start row + half of excess height
            double topOffset = 0;
            for (int row = 0; row < targetArea.StartRow; row++)
                topOffset += sheet.Cells.GetRowHeightPixel(row);
            topOffset += excessHeight / 2;

            // Set position (cast to int)
            picture.Left = (int)leftOffset;
            picture.Top = (int)topOffset;

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
