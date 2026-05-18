using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertAndMovePictureApp
{
    class InsertAndMovePicture
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be linked
                string imagePath = "sample.jpg";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Add a picture whose upper‑left corner is linked to cell H12 (row 11, column 7)
                int pictureIndex = worksheet.Pictures.Add(11, 7, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Verify initial position
                Console.WriteLine($"Initial position - UpperLeftRow: {picture.UpperLeftRow}, UpperLeftColumn: {picture.UpperLeftColumn}");

                // Move the picture so its upper‑left corner is now at cell I13 (row 12, column 8)
                picture.Move(12, 8);

                // Verify new position
                Console.WriteLine($"After move - UpperLeftRow: {picture.UpperLeftRow}, UpperLeftColumn: {picture.UpperLeftColumn}");

                // Save the workbook
                string resultPath = "Result.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(resultPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}