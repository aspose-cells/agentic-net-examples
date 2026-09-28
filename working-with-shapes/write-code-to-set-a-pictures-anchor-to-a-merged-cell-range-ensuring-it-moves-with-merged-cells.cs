// Title: How to anchor a picture to a merged cell range and make it move and resize with the cells using Aspose.Cells for .NET
// AI Prompts: Insert a PNG image into a worksheet, anchor it to a merged range B2:D4, and set its Placement to MoveAndSize with Aspose.Cells. | Compute the total pixel width and height of a merged cell block and apply those dimensions to a Picture object. | Resize a picture so it exactly fills a merged cell area and automatically moves and scales when the merged cells are resized.
// Common Searches: Aspose.Cells C# anchor image to merged cells B2:D4 | Set picture placement to MoveAndSize for a merged range in a .NET workbook | Calculate pixel size of merged cells using Aspose.Cells | Resize picture to fit merged cell block in Aspose.Cells for C#
// Tags: picture placement moveandsize merged cells | merged cell pixel size calculation aspnet | add png picture to merged cell block aspose.cells | resize picture to fill merged area c# | image anchoring merged block aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells B2:D4, inserts a PNG picture anchored at the start of the merged range, sets its Placement to MoveAndSize, calculates the combined pixel dimensions of the merged cells, resizes the picture to fill the merged area, and saves the file as MergedCellPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define merged range B2:D4 (zero‑based indices)
            int startRow = 1;      // B2
            int startColumn = 1;   // column B
            int totalRows = 3;     // rows 2‑4
            int totalColumns = 3;  // columns B‑D

            // Merge the cells
            sheet.Cells.Merge(startRow, startColumn, totalRows, totalColumns);

            // Path to the image file
            string imagePath = "sample.png";

            // Ensure the image file exists
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Add the picture to the worksheet
            using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                int pictureIndex = sheet.Pictures.Add(startRow, startColumn, imgStream);
                Picture picture = sheet.Pictures[pictureIndex];

                // Make the picture move and size with the merged cells
                picture.Placement = PlacementType.MoveAndSize;

                // Calculate total width and height of the merged range in pixels
                double totalWidth = 0;
                for (int col = startColumn; col < startColumn + totalColumns; col++)
                    totalWidth += sheet.Cells.GetColumnWidthPixel(col);

                double totalHeight = 0;
                for (int row = startRow; row < startRow + totalRows; row++)
                    totalHeight += sheet.Cells.GetRowHeightPixel(row);

                // Resize the picture to fit the merged area
                picture.Width = (int)totalWidth;
                picture.Height = (int)totalHeight;
            }

            // Save the workbook
            string outputPath = "MergedCellPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
