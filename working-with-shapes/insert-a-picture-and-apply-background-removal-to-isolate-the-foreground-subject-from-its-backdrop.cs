using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureBackgroundRemoval
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

                // Path to the source image
                string imagePath = "input.jpg";

                // Verify that the image file exists before adding it to the worksheet
                if (File.Exists(imagePath))
                {
                    // Add the picture to the worksheet (row 5, column 2 as an example)
                    int pictureIndex = worksheet.Pictures.Add(5, 2, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Make pure white (RGB 255,255,255) transparent
                    CellsColor transparentColor = workbook.CreateCellsColor();
                    transparentColor.Color = Color.White;
                    picture.FormatPicture.TransparentColor = transparentColor;

                    // Optionally adjust picture size
                    picture.Width = 300;   // width in pixels
                    picture.Height = 200;  // height in pixels
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                }

                // Save the workbook with the modified picture
                workbook.Save("OutputWithTransparentBackground.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}