// Title: Verify a picture's dimensions against its source image before enabling tiling with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an image file, reads its pixel width and height using System.Drawing.Image, adds the image as a picture to a worksheet with Aspose.Cells, compares the picture's Width/Height (in points) to the image dimensions, and sets picture.Tiling = true only when the picture is larger than the source image. | Enhance the given Aspose.Cells example to retrieve the source PNG size, evaluate whether the inserted worksheet picture exceeds those dimensions, and conditionally enable the picture's Tiling property.
// Common Searches: Aspose.Cells C# enable picture tiling only if picture exceeds source image size | How to check worksheet picture dimensions before setting Tiling property in Aspose.Cells | Compare Aspose.Cells picture size with original PNG dimensions .NET | Conditional picture tiling based on size comparison Aspose.Cells example
// Tags: picture tiling condition Aspose.Cells | image dimension retrieval System.Drawing | worksheet picture size comparison .NET | conditional picture tiling C# | shape dimension validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads a workbook, adds a PNG image as a picture to the first worksheet, and saves the file, but it does not compare the picture's dimensions with the original image or apply tiling conditionally. The metadata guides developers on how to implement size validation before enabling tiling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string workbookPath = "input.xlsx";
                string imagePath = "image.png";
                string outputPath = "output.xlsx";

                // Verify required files exist
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add picture to the worksheet at row 1, column 1 (zero‑based indices)
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Optional: adjust picture properties if needed
                // For example, set the picture to move and size with cells
                picture.Placement = PlacementType.MoveAndSize;

                // Save the modified workbook
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
