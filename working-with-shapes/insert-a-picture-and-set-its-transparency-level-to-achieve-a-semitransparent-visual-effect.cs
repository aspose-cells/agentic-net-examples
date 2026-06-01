using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureTransparencyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image to be inserted
                string imagePath = "sampleImage.png";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add the picture to the worksheet (row 2, column 2)
                    int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Set the picture's transparency (0.0 = opaque, 1.0 = fully transparent)
                    picture.FormatPicture.Transparency = 0.5; // 50% transparent
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook to verify the transparency effect
                string outputPath = "PictureTransparencyDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}