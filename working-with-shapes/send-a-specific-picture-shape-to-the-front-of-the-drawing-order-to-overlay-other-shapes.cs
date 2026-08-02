using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two rectangle shapes to demonstrate z‑order
            Shape rect1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
            Shape rect2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Path to the picture file (replace with a valid image path)
            string picturePath = "sample.png";

            // Ensure the picture file exists before adding it
            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Image file not found: {picturePath}");
                return;
            }

            // Add the picture to the worksheet; Add returns the picture index
            int pictureIndex = worksheet.Pictures.Add(2, 2, picturePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Bring the picture to the front so it overlays the rectangles
            picture.ToFrontOrBack(1); // Positive value moves the shape forward

            // Save the workbook
            string outputPath = "PictureFrontDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}