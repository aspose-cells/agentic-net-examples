using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddPictureWithAspectRatioLockApp
{
    class AddPictureWithAspectRatioLock
    {
        static void Main()
        {
            try
            {
                // Path to the image file
                string imagePath = "image.jpg";

                // Ensure the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the picture to the worksheet
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Lock the aspect ratio to maintain proportions when resized
                picture.IsAspectRatioLocked = true;

                // Save the workbook
                string outputPath = "PictureWithAspectRatioLock.xlsx";
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