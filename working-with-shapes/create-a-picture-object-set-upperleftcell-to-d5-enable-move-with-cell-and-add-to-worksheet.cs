using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the image file
                string imagePath = "sample.jpg";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the picture to cell D5 (row index 4, column index 3)
                int picIndex = worksheet.Pictures.Add(4, 3, imagePath);
                Picture picture = worksheet.Pictures[picIndex];

                // Set the picture to move with the cell (but not resize)
                picture.Placement = PlacementType.Move;

                // Ensure the picture is positioned at D5
                picture.Move(4, 3);

                // Save the workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}