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
            // Path to the image file to be inserted
            const string imagePath = "image.jpg";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
                return;
            }

            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the picture to cell B2 (row index 1, column index 1)
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Lock aspect ratio so resizing maintains original proportions
            picture.IsAspectRatioLocked = true;

            // Set picture height to 200 points; width adjusts automatically
            picture.HeightPt = 200;

            // Save the workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}