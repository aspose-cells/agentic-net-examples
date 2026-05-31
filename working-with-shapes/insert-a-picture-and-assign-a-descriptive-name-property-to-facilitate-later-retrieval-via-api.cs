using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string imagePath = "sampleImage.png";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add picture to the worksheet (rows 0‑4, columns 0‑4)
                    int pictureIndex = worksheet.Pictures.Add(0, 0, 4, 4, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];
                    picture.Name = "CompanyLogo";
                    picture.AlternativeText = "Company logo displayed in the report";
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "PictureWithName.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}