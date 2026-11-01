// Title: Insert a PNG picture into an Excel worksheet and proportionally scale it to fit a B2:D10 cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a PNG image into a worksheet, compute the pixel dimensions of cells B2 to D10, scale the image while preserving its aspect ratio, and place it at the top‑left corner using Aspose.Cells in C#. | Resize and position a picture so that it fits exactly within a specified cell range in an Excel file with Aspose.Cells, ensuring proportional scaling based on the range's pixel size.
// Common Searches: Aspose.Cells C# insert image and fit it into a specific cell range preserving aspect ratio | calculate pixel width of Excel cell range with Aspose.Cells | scale picture to match Excel cells B2:D10 using Aspose.Cells .NET | position picture at top left of a cell block in Aspose.Cells workbook
// Tags: insert picture Aspose.Cells C# | scale image to cell range Aspose.Cells | calculate cell range pixel dimensions Aspose.Cells | preserve aspect ratio picture Aspose.Cells | position picture top left cell range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Inserts a PNG image into a new workbook, determines the pixel size of the B2:D10 range, scales the image proportionally, positions it at the range's top‑left corner, and saves the file as output.xlsx.
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

            // Convert the cell addresses to zero‑based row/column indexes using Cells collection
            Cell start = sheet.Cells[startCell];
            int startRow = start.Row;
            int startColumn = start.Column;

            Cell end = sheet.Cells[endCell];
            int endRow = end.Row;
            int endColumn = end.Column;

            // Calculate the pixel dimensions of the target range
            int targetWidth = 0;
            for (int col = startColumn; col <= endColumn; col++)
            {
                targetWidth += sheet.Cells.GetColumnWidthPixel(col);
            }

            int targetHeight = 0;
            for (int row = startRow; row <= endRow; row++)
            {
                targetHeight += sheet.Cells.GetRowHeightPixel(row);
            }

            // Calculate the top‑left pixel offset of the target range within the worksheet
            int offsetLeft = 0;
            for (int col = 0; col < startColumn; col++)
            {
                offsetLeft += sheet.Cells.GetColumnWidthPixel(col);
            }

            int offsetTop = 0;
            for (int row = 0; row < startRow; row++)
            {
                offsetTop += sheet.Cells.GetRowHeightPixel(row);
            }

            // Path to the image file to be inserted
            string imagePath = @"C:\Images\sample.png";

            // Ensure the image file exists before attempting to insert
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at the top‑left corner of the worksheet (temporary position)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Original dimensions of the picture (in pixels)
            int originalWidth = picture.Width;
            int originalHeight = picture.Height;

            // Determine the scaling factor to fit the picture proportionally inside the target range
            double widthScale = (double)targetWidth / originalWidth;
            double heightScale = (double)targetHeight / originalHeight;
            double scale = Math.Min(widthScale, heightScale); // preserve aspect ratio

            // Apply the scaling factor
            picture.Width = (int)(originalWidth * scale);
            picture.Height = (int)(originalHeight * scale);

            // Position the picture at the calculated offset (top‑left of the target range)
            picture.Left = offsetLeft;
            picture.Top = offsetTop;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
