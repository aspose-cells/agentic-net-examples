// Title: How to proportionally resize an inserted picture to fit a specific cell range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# resize picture to fit within a cell range while maintaining aspect ratio | scale image to match Excel cells C5:E10 with Aspose.Cells .NET API | fit a PNG shape into a target range and keep its proportions using Aspose.Cells | calculate column and row pixel dimensions for shape scaling in Aspose.Cells | set picture placement to MoveAndSize so it moves with cells in Aspose.Cells C#
// Tags: picture scaling based on cell range dimensions Aspose.Cells | maintain image aspect ratio when resizing shape Excel | compute column width in pixels Aspose.Cells | compute row height in pixels Aspose.Cells | picture placement MoveAndSize Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a PNG picture, determines the combined pixel width and height of a target cell range (e.g., C5:E10), calculates the minimum scaling factor to preserve the picture's original aspect ratio, resizes the picture accordingly, aligns its upper‑left corner with the first cell of the range, sets the placement to MoveAndSize so the shape moves and resizes with the cells, and saves the workbook as ResizedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image that will be inserted.
            // Replace with a valid image file on your machine.
            string imagePath = "sample_image.png";

            // Ensure the image exists; create a dummy PNG if it does not.
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file \"{imagePath}\" not found. Creating a dummy 1x1 red PNG.");
                CreateDummyPng(imagePath);
            }

            // -----------------------------------------------------------------
            // 1. Create a workbook and get the first worksheet.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 2. Insert the picture into the worksheet (initial position does not matter).
            // -----------------------------------------------------------------
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // -----------------------------------------------------------------
            // 3. Define the target cell range that the picture should fit into.
            //    Example: range C5:E10
            // -----------------------------------------------------------------
            string targetRangeAddress = "C5:E10";
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange(targetRangeAddress);

            // -----------------------------------------------------------------
            // 4. Calculate the total width (in pixels) of the columns that belong to the range.
            // -----------------------------------------------------------------
            int startCol = targetRange.FirstColumn;
            int endCol   = targetRange.FirstColumn + targetRange.ColumnCount - 1;
            double totalWidthPixels = 0;
            for (int col = startCol; col <= endCol; col++)
            {
                totalWidthPixels += sheet.Cells.GetColumnWidthPixel(col);
            }

            // -----------------------------------------------------------------
            // 5. Calculate the total height (in pixels) of the rows that belong to the range.
            // -----------------------------------------------------------------
            int startRow = targetRange.FirstRow;
            int endRow   = targetRange.FirstRow + targetRange.RowCount - 1;
            double totalHeightPixels = 0;
            for (int row = startRow; row <= endRow; row++)
            {
                totalHeightPixels += sheet.Cells.GetRowHeightPixel(row);
            }

            // -----------------------------------------------------------------
            // 6. Determine the scaling factor that preserves the picture's aspect ratio.
            //    The picture will be scaled so that it fits completely inside the target range.
            // -----------------------------------------------------------------
            double originalWidthPixels  = picture.OriginalWidth;
            double originalHeightPixels = picture.OriginalHeight;

            double widthScale  = totalWidthPixels  / originalWidthPixels;
            double heightScale = totalHeightPixels / originalHeightPixels;
            double scaleFactor = Math.Min(widthScale, heightScale); // keep aspect ratio

            // -----------------------------------------------------------------
            // 7. Apply the new size to the picture.
            //    Width and Height properties are expressed in points (1 point = 1/72 inch).
            //    Cast to int if the API expects integer values.
            // -----------------------------------------------------------------
            picture.Width  = (int)(originalWidthPixels  * scaleFactor);
            picture.Height = (int)(originalHeightPixels * scaleFactor);

            // -----------------------------------------------------------------
            // 8. Position the picture so that its upper‑left corner aligns with the
            //    upper‑left cell of the target range and make it move/size with the cells.
            // -----------------------------------------------------------------
            picture.UpperLeftRow    = startRow;
            picture.UpperLeftColumn = startCol;
            picture.Placement       = PlacementType.MoveAndSize;

            // -----------------------------------------------------------------
            // 9. Save the workbook.
            // -----------------------------------------------------------------
            string outputPath = "ResizedShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Creates a minimal red PNG file from a base64 string (avoids System.Drawing dependency).
    private static void CreateDummyPng(string path)
    {
        // 1x1 red pixel PNG (base64 encoded)
        const string base64RedPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64RedPng);
        File.WriteAllBytes(path, pngBytes);
    }
}
