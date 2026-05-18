using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCropDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the image file to be inserted
                string imagePath = "sample.jpg";

                // Verify that the image file exists before attempting to add it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add the picture to cell B2 (row index 1, column index 1)
                int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Access the picture's format object to apply cropping
                MsoFormatPicture format = picture.FormatPicture;
                format.LeftCrop = 0.25;   // Crop 25% from the left
                format.RightCrop = 0.25;  // Crop 25% from the right
                format.TopCrop = 0.25;    // Crop 25% from the top
                format.BottomCrop = 0.25; // Crop 25% from the bottom

                // Save the workbook
                string outputPath = "CroppedPictureDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}