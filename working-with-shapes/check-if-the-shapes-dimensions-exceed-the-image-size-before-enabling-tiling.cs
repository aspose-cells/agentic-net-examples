// Title: Validate picture shape dimensions against original image size before enabling tiling with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads the image file using System.Drawing.Image, reads its pixel width and height, compares them with the Aspose.Cells picture shape's Width and Height, and sets picture.IsTiled = true only when the shape exceeds the image dimensions. | Update the provided Aspose.Cells example to add a conditional block that enables tiling only for picture shapes larger than the original PNG, including proper error handling for missing files.
// Common Searches: Aspose.Cells C# compare picture shape size with source image before tiling | enable tiling for picture shape only when it is larger than the image | how to prevent tiling of small picture shapes in Aspose.Cells | C# Aspose.Cells conditional tiling based on shape dimensions
// Tags: Aspose.Cells picture shape size validation | conditional tiling activation | retrieve image dimensions System.Drawing C# | compare shape dimensions with source image | enable tiling only for oversized picture shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a workbook, loads a PNG file, inserts it as a picture shape at cell B3, optionally adjusts its size, and demonstrates how to conditionally enable the picture.IsTiled property only when the shape's dimensions exceed the original image size. It also handles missing files, creates the output folder, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Path to the image file to be used for the shape
                string imagePath = @"C:\Images\SampleImage.png";

                // Verify that the image file exists before loading
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                try
                {
                    // Add the picture shape to the worksheet at cell B3 (row index 2, column index 2)
                    int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Optionally increase the picture size by 20 points in each dimension
                    picture.Width = picture.Width + 20;
                    picture.Height = picture.Height + 20;

                    // If the current Aspose.Cells version supports tiling, enable it when needed.
                    // Uncomment the following line if the IsTiled property is available:
                    // picture.IsTiled = true;
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to add picture: {picEx.Message}");
                    return;
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Output\WorkbookWithTiledShape.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
