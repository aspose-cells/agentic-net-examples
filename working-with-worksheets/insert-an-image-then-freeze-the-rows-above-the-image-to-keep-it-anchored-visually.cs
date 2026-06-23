using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageFreezeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file (ensure the file exists at this location)
                string imagePath = "image.jpg";

                // Verify that the image file exists before attempting to add it
                if (!File.Exists(imagePath))
                {
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
                }

                // Define the cell where the image's top‑left corner will be placed
                // Row 5 (zero‑based index 4) and column B (zero‑based index 1)
                int topRow = 4;
                int leftColumn = 1;

                // Add the picture to the worksheet
                int pictureIndex = worksheet.Pictures.Add(topRow, leftColumn, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Optional: adjust picture size or placement if needed
                // picture.Width = 200;
                // picture.Height = 150;

                // Freeze all rows above the image so the picture stays visible while scrolling
                // FreezePanes(row, column, freezedRows, freezedColumns)
                worksheet.FreezePanes(topRow, 0, topRow, 0);

                // Save the workbook
                string outputPath = "ImageWithFrozenRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (CellsException ex)
            {
                Console.Error.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}