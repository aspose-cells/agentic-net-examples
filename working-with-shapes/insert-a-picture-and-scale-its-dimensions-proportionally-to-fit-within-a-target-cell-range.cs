// Title: Insert a picture into an Excel worksheet and proportionally scale it to fit a target cell range using Aspose.Cells for .NET
// AI Prompts: Insert an image at cell B2 and resize it proportionally so it fits within the range B2:D10 with Aspose.Cells for .NET. | Compute the pixel width and height of a cell range and apply a uniform scaling factor to a picture to preserve its aspect ratio in Aspose.Cells. | After scaling a picture to a cell range, center it horizontally and vertically inside the range using Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to fit an image inside a specific Excel cell range while keeping aspect ratio | scale picture to match Excel range dimensions using Aspose.Cells C# | center inserted image within B2:D10 range in Aspose.Cells workbook
// Tags: insert picture into worksheet Aspose.Cells .NET | proportional image scaling Excel range | calculate cell range pixel size Aspose.Cells | center picture within Excel cells Aspose.Cells | maintain aspect ratio picture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, defines a target range (B2:D10), loads an image, inserts it at the range's upper‑left cell, calculates the total pixel width and height of the range, determines the smallest scaling factor to keep the image's aspect ratio, resizes the picture accordingly, optionally centers it within the range, and saves the result to output.xlsx.
class InsertAndScalePicture
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range where the picture should fit (e.g., B2:D10)
            string startCell = "B2";
            string endCell   = "D10";

            // Convert cell names to row/column indexes using the Cells collection
            int startRow = sheet.Cells[startCell].Row;
            int startCol = sheet.Cells[startCell].Column;
            int endRow   = sheet.Cells[endCell].Row;
            int endCol   = sheet.Cells[endCell].Column;

            // Path to the image file to be inserted
            string imagePath = "sample.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Insert the picture at the upper‑left cell of the target range
            int pictureIndex = sheet.Pictures.Add(startRow, startCol, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // ----- Compute target dimensions in pixels -----
            double targetWidthPx = 0;
            for (int col = startCol; col <= endCol; col++)
                targetWidthPx += sheet.Cells.GetColumnWidthPixel(col);

            double targetHeightPx = 0;
            for (int row = startRow; row <= endRow; row++)
                targetHeightPx += sheet.Cells.GetRowHeightPixel(row);

            // ----- Compute scaling factor to keep aspect ratio -----
            double widthScale  = targetWidthPx  / picture.Width;
            double heightScale = targetHeightPx / picture.Height;
            double scaleFactor = Math.Min(widthScale, heightScale);

            // Apply proportional scaling
            picture.Width  = (int)(picture.Width  * scaleFactor);
            picture.Height = (int)(picture.Height * scaleFactor);

            // Optional: center the picture within the target range
            double remainingWidth  = targetWidthPx  - picture.Width;
            double remainingHeight = targetHeightPx - picture.Height;
            picture.Left = (int)(remainingWidth  / 2);
            picture.Top  = (int)(remainingHeight / 2);

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
