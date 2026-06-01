using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains one default sheet named "Sheet1")
                Workbook workbook = new Workbook();

                // Add additional worksheets with unique names
                workbook.Worksheets.Add("Sheet1_Added");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Path to the image that will be linked (local file or URL)
                string imagePath = "image.jpg";

                // Verify that the image file exists before using it
                bool imageExists = File.Exists(imagePath);
                if (!imageExists)
                {
                    Console.WriteLine($"Warning: Image file not found: {imagePath}");
                }

                // Insert a linked picture at cell A1 of each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Cells["A1"].PutValue("Linked picture below:");

                    if (imageExists)
                    {
                        // row, column, height, width (pixels)
                        sheet.Shapes.AddLinkedPicture(0, 0, 100, 100, imagePath);
                    }
                }

                // Save the workbook
                string outputPath = "output_with_linked_pictures.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}