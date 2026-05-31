using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureRotationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "image.jpg";

            // Ensure the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet at a specific cell range (A1 to B5)
            // Using the Pictures collection overload that takes a file name
            int pictureIndex = worksheet.Pictures.Add(0, 0, 4, 1, imagePath);

            // Retrieve the inserted picture object
            Picture picture = worksheet.Pictures[pictureIndex];

            // Rotate the picture 90 degrees clockwise (positive angle)
            picture.RotationAngle = 90;

            // Save the workbook
            workbook.Save("PictureRotated.xlsx");

            Console.WriteLine("Picture added and rotated successfully.");
        }
    }
}