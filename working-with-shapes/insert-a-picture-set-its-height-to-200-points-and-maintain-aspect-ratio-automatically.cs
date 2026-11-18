// Title: Insert an image into an Excel worksheet, set its height to 200 points, and lock aspect ratio using Aspose.Cells for .NET (C#)
// Common Searches: how to add a PNG picture to an Excel worksheet with Aspose.Cells and keep its aspect ratio | C# Aspose.Cells set picture height in points while preserving width proportion | example of using picture.IsLockAspectRatio true to scale images in Aspose.Cells | insert image at cell A1 and fix height to 200 points using Aspose.Cells for .NET | save workbook after adding a resized picture with Aspose.Cells C#
// Tags: add picture to worksheet Aspose.Cells C# | set picture height points Aspose.Cells | lock aspect ratio picture Aspose.Cells | image scaling Excel .NET | insert PNG image Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# program creates a new workbook, adds a PNG image to the first worksheet, sets the picture's height to 200 points, enables IsLockAspectRatio to preserve the original width proportionally, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "image.png";

            // Ensure the image file exists before adding it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet; Add returns the picture index
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the picture object using the index
            Picture picture = sheet.Pictures[pictureIndex];

            // Set the picture height to 200 points
            picture.Height = 200;

            // Lock the aspect ratio so width adjusts automatically
            picture.IsLockAspectRatio = true;

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
