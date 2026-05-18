using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertPictureInCellApp
{
    class InsertPictureInCell
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell (zero‑based indices). Example: cell D3 -> row 2, column 3
                int targetRow = 2;      // Row index for D3
                int targetColumn = 3;   // Column index for D3

                // Local image file path
                string imagePath = @"C:\Images\sample.png";

                // Verify that the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Add the picture to the worksheet at the specified cell's top‑left corner
                int pictureIndex = worksheet.Pictures.Add(targetRow, targetColumn, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Place the picture inside the cell (instead of floating over cells)
                picture.IsPlacedInCell = true;

                // Save the workbook
                string outputPath = "OutputWithPicture.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}