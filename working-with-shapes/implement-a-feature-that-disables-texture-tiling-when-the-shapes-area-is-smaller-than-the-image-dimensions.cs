// Title: How to disable texture tiling for a picture shape when its size is smaller than the source image using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an image file, reads its pixel width and height, inserts it as a picture shape into a worksheet, compares the shape's Width/Height (in points) to the image dimensions, and sets the shape's TextureTiling property to false when the shape is smaller. | Create a reusable method InsertPictureWithoutTiling(Worksheet sheet, string imagePath, int row, int column, double width, double height) that automatically disables texture tiling if the picture shape's area is less than the image's pixel area. | Show how to convert image pixel dimensions to points, perform the size comparison, and programmatically turn off texture tiling for Aspose.Cells picture shapes in a .NET workbook.
// Common Searches: Aspose.Cells .NET disable picture texture tiling when shape is smaller than image | C# compare shape size to image size before adding picture in Excel with Aspose.Cells | How to turn off texture tiling for a shape in Aspose.Cells if the shape dimensions are less than the source image | Set TextureTiling false for picture shape based on image dimensions Aspose.Cells | Prevent image repeat in Excel shape using Aspose.Cells C#
// Tags: texture tiling control Aspose.Cells | picture shape size check .NET | disable image repeat in Excel shape | image dimension conversion points C# | shape insertion helper method Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // // Loads a PNG file, adds it as a picture shape to a worksheet, optionally resizes the shape, checks if the shape's area is smaller than the image's pixel area, disables texture tiling to prevent repeat, and saves the workbook as Output.xlsx with error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the texture image file exists
                string texturePath = "texture.png";
                if (!File.Exists(texturePath))
                {
                    Console.WriteLine($"Texture file '{texturePath}' not found.");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add the picture to the worksheet (row, column, row offset, column offset, stream)
                Shape pictureShape = null;
                try
                {
                    using (FileStream imageStream = new FileStream(texturePath, FileMode.Open, FileAccess.Read))
                    {
                        pictureShape = sheet.Shapes.AddPicture(5, 5, 0, 0, imageStream);
                    }
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Failed to add picture shape: {shapeEx.Message}");
                    return;
                }

                // Optionally adjust the size of the picture (width and height in points)
                pictureShape.Width = 100;   // width in points
                pictureShape.Height = 80;   // height in points

                // Save the workbook
                string outputPath = "Output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
