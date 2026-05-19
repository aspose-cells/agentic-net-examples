using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
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

                // Path to the image file
                string imagePath = "sampleImage.png";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    try
                    {
                        // Add the picture to the worksheet (row 2, column 2)
                        int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                        Picture picture = worksheet.Pictures[pictureIndex];

                        // Send the picture to the back of the z-order (0 = back, 1 = front)
                        picture.ToFrontOrBack(0);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to insert picture: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "PictureBackDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}