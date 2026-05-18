using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCameraDemo
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

                // Populate sample data that will be captured by the camera
                worksheet.Cells["A1"].Value = "Header 1";
                worksheet.Cells["B1"].Value = "Header 2";
                worksheet.Cells["A2"].Value = "Row 1, Col 1";
                worksheet.Cells["B2"].Value = "Row 1, Col 2";
                worksheet.Cells["A3"].Value = "Row 2, Col 1";
                worksheet.Cells["B3"].Value = "Row 2, Col 2";

                // Get the Pictures collection from the worksheet
                PictureCollection pictures = worksheet.Pictures;

                // Create a camera picture that captures the range A1:B3.
                // The picture will be placed with its top‑left corner at row 5, column 1.
                int pictureIndex = pictures.Camera(5, 1, "A1:B3");

                // Retrieve the created picture and export it as an image file
                Picture cameraPicture = pictures[pictureIndex];
                using (FileStream imageStream = new FileStream("CameraCapture.png", FileMode.Create, FileAccess.Write))
                {
                    cameraPicture.ToImage(imageStream, ImageType.Png);
                }

                // Save the workbook containing the camera picture
                workbook.Save("CameraDemo.xlsx");

                Console.WriteLine($"Camera picture added at index {pictureIndex}");
                Console.WriteLine("Workbook and image saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}