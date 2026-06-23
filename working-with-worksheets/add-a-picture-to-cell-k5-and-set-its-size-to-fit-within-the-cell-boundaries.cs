using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddPictureToCellApp
{
    class AddPictureToCell
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string imagePath = "image.png";

                // Ensure the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add a picture to cell K5 (row 4, column 10) within cell boundaries
                    int pictureIndex = worksheet.Pictures.Add(4, 10, 4, 10, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];
                    picture.IsPlacedInCell = true; // anchor the picture to the cell
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                }

                // Save the workbook
                workbook.Save("Output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}