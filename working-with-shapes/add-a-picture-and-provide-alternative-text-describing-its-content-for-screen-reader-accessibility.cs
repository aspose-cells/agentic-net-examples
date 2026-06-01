using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class AddPictureWithAltText
    {
        // Entry point required for console application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the image file
            string imagePath = "image.png";

            // Ensure the image file exists before adding it to the worksheet
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a picture at row 2, column 2 (zero‑based indices)
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Set alternative text for screen readers / accessibility
            picture.AlternativeText = "A sample picture showing a sunrise over mountains";

            // Save the workbook to a file
            workbook.Save("PictureWithAltText.xlsx");
        }
    }
}